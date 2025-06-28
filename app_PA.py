import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO

def read_csv_robust(uploaded_file):
    try:
        return pd.read_csv(uploaded_file, encoding="ISO-8859-1")
    except Exception:
        try:
            return pd.read_csv(uploaded_file, encoding="utf-8")
        except Exception as e2:
            st.error(f"Unable to load CSV file. Try saving it again via Excel as 'CSV'. Error: {e2}")
            return None

def standardize_columns(df, region="auto"):
    df = df.copy()
    if region == "auto":
        if "List Number" in df.columns:
            region = "PA"
        else:
            region = "KC"
    if region == "PA":
        df['Address'] = (
            df.get("House # (1)", '').fillna('').astype(str).str.strip() + " " +
            df.get("Street Name", '').fillna('').astype(str).str.strip() + ", " +
            df.get("City (Post Office)", '').fillna('').astype(str).str.strip() + ", PA " +
            df.get("Zip Code", '').fillna('').astype(str).str.strip()
        )
        column_mapping = {
            "List Price": "List Price",
            "Sold Price": "Sale Price",
            "Total Bedrooms": "Bedrooms",
            "Zip Code": "Zip",
            "County": "County",
            "City (Post Office)": "City",
            "Aprx Total SF": "Total Finished SF",
            "Develop/Subdivision": "Sub",
            "List Number": "MLS #",
            "List Date": "List Dt",
            "Full Baths": "Full Baths",
            "Status": "Status",
            "Area": "Area",
            "MLS Area": "MLS Area",
            "Close Dt": "Close Dt",
            "Est Total Tax": "Property Tax"
        }
        for original, new_name in column_mapping.items():
            if original in df.columns:
                df[new_name] = df[original]
        for tax_name in ["Tax", "Taxes", "Yearly Tax", "Annual Tax", "Property taxes"]:
            if tax_name in df.columns and "Property Tax" not in df.columns:
                df["Property Tax"] = df[tax_name]
        if "Zillow" not in df.columns:
            df["Zillow"] = ""
    if "Zip" in df.columns:
        df["Zip"] = df["Zip"].astype(str).str.strip().str.zfill(5)
    return df

def build_summary_table(df_listings, df_comps, sort_by):
    if sort_by == "ALL":
        groups = ["ALL"]
        listings_grouped = [df_listings]
        comps_grouped = [df_comps]
    else:
        unique_vals = df_listings[sort_by].dropna().unique()
        order = list(pd.Series(unique_vals).drop_duplicates())
        groups = []
        for g in order:
            groups.append(g)
        for g in df_comps[sort_by].dropna().unique():
            if g not in groups:
                groups.append(g)
        listings_grouped = [df_listings[df_listings[sort_by] == group] for group in groups]
        comps_grouped = [df_comps[df_comps[sort_by] == group] for group in groups]
    rows = []
    for group, df_l, df_c in zip(groups, listings_grouped, comps_grouped):
        n_listings = len(df_l)
        avg_listing = df_l["List Price"].mean() if n_listings > 0 else float('nan')
        n_sold = len(df_c)
        avg_sold = df_c["Sale Price"].mean() if n_sold > 0 else float('nan')
        diff = avg_sold - avg_listing if pd.notnull(avg_sold) and pd.notnull(avg_listing) else float('nan')
        pct = (diff / avg_listing * 100) if pd.notnull(diff) and avg_listing and avg_listing != 0 else float('nan')
        rows.append({
            sort_by: group,
            "# Listings": n_listings,
            "Avg Listing Price": avg_listing,
            "# Sold": n_sold,
            "Avg Sold Price": avg_sold,
            "$ Diff": diff,
            "% Diff": pct
        })
    df_summary = pd.DataFrame(rows)
    df_summary = df_summary.sort_values(by="# Listings", ascending=False)
    return df_summary, df_summary[sort_by].tolist()

def build_zillow_hyperlink(address):
    if pd.isnull(address) or address.strip() == "":
        return ""
    url = f"https://www.zillow.com/homes/{address.replace(' ', '-').replace(',', '')}_rb/"
    return f'=HYPERLINK("{url}", "Zillow")'

def export_full_analysis(summary_table, df_props, flip_row, comps_display, criteria_dict):
    summary_tab = summary_table.copy()
    focused_tab = df_props.copy()  # This should include Zillow HYPERLINK column
    flip_tab = flip_row.copy()
    comps_tab = comps_display.copy()
    # Ensure 'Zillow' column is present in flip_tab for export
    if "Address" in flip_tab.columns and "Zillow" not in flip_tab.columns:
        flip_tab["Zillow"] = flip_tab["Address"].apply(build_zillow_hyperlink)
    # Build Flip+Comps table
    if "Zillow" in comps_tab.columns:
        flip_and_comps = pd.concat([
            flip_tab.reset_index(drop=True),
            pd.DataFrame([[""] * flip_tab.shape[1]], columns=flip_tab.columns),
            comps_tab.reset_index(drop=True)
        ], ignore_index=True)
    else:
        # Add empty 'Zillow' col to comps if not present
        comps_tab["Zillow"] = ""
        flip_and_comps = pd.concat([
            flip_tab.reset_index(drop=True),
            pd.DataFrame([[""] * flip_tab.shape[1]], columns=flip_tab.columns),
            comps_tab.reset_index(drop=True)
        ], ignore_index=True)
    crit_tab = pd.DataFrame([criteria_dict])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_tab.to_excel(writer, index=False, sheet_name='Market Summary')
        focused_tab.to_excel(writer, index=False, sheet_name='Focused Flips')
        flip_and_comps.to_excel(writer, index=False, sheet_name='Flip+Comps')
        crit_tab.to_excel(writer, index=False, sheet_name='Comps Criteria')
        # Focused Flips hyperlinks
        workbook = writer.book
        worksheet = writer.sheets['Focused Flips']
        if "Zillow" in focused_tab.columns:
            zillow_col_idx = focused_tab.columns.get_loc("Zillow")
            hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': 1})
            for rownum in range(1, len(focused_tab) + 1):
                formula = focused_tab.iloc[rownum-1]['Zillow']
                if formula and str(formula).startswith('=HYPERLINK'):
                    worksheet.write_formula(rownum, zillow_col_idx, formula, hyperlink_format)
        # Flip+Comps hyperlinks (flip row + comps)
        worksheet_fc = writer.sheets['Flip+Comps']
        if "Zillow" in flip_and_comps.columns:
            zillow_col_idx_fc = flip_and_comps.columns.get_loc("Zillow")
            hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': 1})
            for rownum in range(1, len(flip_and_comps) + 1):
                formula = flip_and_comps.iloc[rownum-1]['Zillow']
                if formula and str(formula).startswith('=HYPERLINK'):
                    worksheet_fc.write_formula(rownum, zillow_col_idx_fc, formula, hyperlink_format)
    output.seek(0)
    return output

for key in ["focused_analysis", "df_props", "df_focus", "df_comps_focus"]:
    if key not in st.session_state:
        if "df" in key:
            st.session_state[key] = pd.DataFrame()
        else:
            st.session_state[key] = False

st.set_page_config(page_title="🏠 NewRoof PA Screening Analyzer", layout="wide")
st.markdown("<h1 style='text-align: center; color: teal;'>🏠 NewRoof PA Screening Analyzer</h1>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
df_listings = None
df_comps = None

with col1:
    listings_file = st.file_uploader("Upload Listings CSV", type="csv", key="listings")
    if listings_file:
        df_listings = read_csv_robust(listings_file)
        if df_listings is not None:
            df_listings = standardize_columns(df_listings)
            st.success(f"✅ {len(df_listings):,} rows loaded from Listings file.")

with col2:
    comps_file = st.file_uploader("Upload Comps CSV", type="csv", key="comps")
    if comps_file:
        df_comps = read_csv_robust(comps_file)
        if df_comps is not None:
            df_comps = standardize_columns(df_comps)
            st.success(f"✅ {len(df_comps):,} rows loaded from Comps file.")

if (df_listings is not None) and (df_comps is not None):
    desired_sort_columns = ["Area", "MLS Area", "County", "City", "Sub"]
    available_sort_columns = [col for col in desired_sort_columns if col in df_listings.columns and df_listings[col].notnull().sum() > 0]
    sort_options = ["ALL"] + available_sort_columns

    sort_selection = st.selectbox("Sort/filter by", sort_options)
    summary_table, ordered_groups = build_summary_table(df_listings, df_comps, sort_selection)
    st.subheader("Market Summary by " + sort_selection)
    st.dataframe(summary_table.style.format({
        "Avg Listing Price": "{:,.0f}",
        "Avg Sold Price": "{:,.0f}",
        "$ Diff": "{:,.0f}",
        "% Diff": "{:,.2f}"
    }), use_container_width=True)

    if sort_selection != "ALL":
        visible_areas = summary_table[sort_selection].tolist()
        selected_groups = st.multiselect(f"Select {sort_selection}(s) to focus on:", visible_areas)
        ready_to_run = bool(selected_groups)
    else:
        selected_groups = []
        ready_to_run = True

    st.markdown("### Comps Matching Criteria")
    match_zip = st.checkbox("Same ZIP", value=True)
    match_county = st.checkbox("Same County", value=False)
    match_city = st.checkbox("Same City", value=False)
    match_sub = st.checkbox("Same Sub", value=False)
    match_beds = st.checkbox("Same # Bedrooms", value=True)
    sf_pct = st.slider("± SF Range (%):", 5, 50, 15, 1)

    def match_comps(row, comps):
        cond = pd.Series([True] * len(comps), index=comps.index)
        if match_zip and "Zip" in row and "Zip" in comps.columns:
            zip_val = str(row["Zip"]).strip().zfill(5) if pd.notnull(row["Zip"]) else ""
            comps_zip = comps["Zip"].astype(str).str.strip().str.zfill(5)
            cond &= comps_zip == zip_val
        if match_county and "County" in row and "County" in comps.columns:
            cond &= comps["County"].astype(str).str.strip() == str(row["County"]).strip()
        if match_city and "City" in row and "City" in comps.columns:
            cond &= comps["City"].astype(str).str.strip() == str(row["City"]).strip()
        if match_sub and "Sub" in row and "Sub" in comps.columns:
            cond &= comps["Sub"].astype(str).str.strip() == str(row["Sub"]).strip()
        if match_beds and "Bedrooms" in row and "Bedrooms" in comps.columns:
            try:
                cond &= comps["Bedrooms"].astype(float).round() == float(row["Bedrooms"])
            except Exception:
                cond &= False
        if sf_pct and "Total Finished SF" in row and "Total Finished SF" in comps.columns:
            try:
                sf = float(row["Total Finished SF"])
                pct = sf_pct / 100
                min_sf = sf * (1 - pct)
                max_sf = sf * (1 + pct)
                cond &= comps["Total Finished SF"].astype(float).between(min_sf, max_sf)
            except Exception:
                cond &= False
        return comps[cond]

    if ready_to_run and st.button("▶️ Run Focused Area Analysis", type="primary"):
        st.session_state['focused_analysis'] = True
        if sort_selection == "ALL":
            st.session_state['df_focus'] = df_listings.copy()
            st.session_state['df_comps_focus'] = df_comps.copy()
        else:
            st.session_state['df_focus'] = df_listings[df_listings[sort_selection].isin(selected_groups)].copy()
            st.session_state['df_comps_focus'] = df_comps[df_comps[sort_selection].isin(selected_groups)].copy()
        rows = []
        for idx, row in st.session_state['df_focus'].iterrows():
            comps_matched = match_comps(row, st.session_state['df_comps_focus'])
            if not comps_matched.empty and "Sale Price" in comps_matched.columns:
                valid_prices = pd.to_numeric(comps_matched["Sale Price"], errors='coerce').dropna()
                avg_comp = valid_prices.mean() if not valid_prices.empty else np.nan
            else:
                avg_comp = np.nan
            dollar_diff = avg_comp - row["List Price"] if pd.notnull(avg_comp) and pd.notnull(row["List Price"]) else np.nan
            pct_diff = (dollar_diff / row["List Price"] * 100) if pd.notnull(dollar_diff) and row["List Price"] else np.nan
            rows.append({
                "MLS #": row.get("MLS #", np.nan),
                "Address": row.get("Address", ""),
                "Bedrooms": row.get("Bedrooms", ""),
                "SF": row.get("Total Finished SF", ""),
                "List Price": row.get("List Price", np.nan),
                "Avg Comp Price": avg_comp,
                "Price Diff ($)": dollar_diff,
                "Price Diff (%)": pct_diff,
                "# of Comps": len(comps_matched)
            })
        df_props = pd.DataFrame(rows)
        if not df_props.empty:
            df_props = df_props.sort_values(by="Price Diff (%)", ascending=False).reset_index(drop=True)
            df_props["Rank"] = df_props.index + 1
        st.session_state['df_props'] = df_props

    df_props = st.session_state['df_props']
    df_focus = st.session_state['df_focus']
    df_comps_focus = st.session_state['df_comps_focus']

    if st.session_state['focused_analysis'] and not df_props.empty:
        show_cols = ["Rank", "MLS #", "Address", "Bedrooms", "SF", "List Price", "Avg Comp Price", "Price Diff ($)", "Price Diff (%)", "# of Comps"]
        st.dataframe(
            df_props[show_cols].style.format({
                "SF": "{:,.0f}",
                "List Price": "{:,.0f}",
                "Avg Comp Price": "{:,.0f}",
                "Price Diff ($)": "{:,.0f}",
                "Price Diff (%)": "{:,.1f}%",
            }),
            use_container_width=True
        )

        mls_options = list(df_props["MLS #"])
        selected_mls = st.multiselect("Select MLS #(s) to focus on:", mls_options)

        if not df_focus.empty and not df_comps_focus.empty and selected_mls:
            for mls_num in selected_mls:
                flip_row = df_focus[df_focus["MLS #"] == mls_num]
                full_address = flip_row["Address"].iloc[0] if not flip_row.empty and "Address" in flip_row.columns else ""
                zillow_search_url = f"https://www.zillow.com/homes/{full_address.replace(' ', '-').replace(',', '')}_rb/"
                st.markdown(
                    f"### 📌 Flip Details: MLS {mls_num} "
                    f"<a href='{zillow_search_url}' target='_blank'>"
                    f"<button style='background:#2563eb;color:white;padding:2px 8px;border-radius:5px;border:none;margin-left:8px;'>Zillow Search</button></a>",
                    unsafe_allow_html=True
                )

                flip_cols = [
                    "MLS #", "Status", "Address", "County", "City", "Zip", "Sub",
                    "Bedrooms", "Total Finished SF", "List Price"
                ]
                if "Property Tax" in flip_row.columns:
                    lp_idx = flip_cols.index("List Price")
                    flip_cols.insert(lp_idx + 1, "Property Tax")
                flip_cols += ["List Dt"]
                # Ensure Zillow in flip_cols for export
                if "Zillow" not in flip_cols:
                    flip_cols.append("Zillow")
                existing_flip_cols = [col for col in flip_cols if col in flip_row.columns]
                flip_display = flip_row[existing_flip_cols].copy()
                if "List Price" in flip_display.columns:
                    flip_display["List Price"] = flip_display["List Price"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                if "Property Tax" in flip_display.columns:
                    flip_display["Property Tax"] = flip_display["Property Tax"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                if "Total Finished SF" in flip_display.columns:
                    flip_display["Total Finished SF"] = flip_display["Total Finished SF"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                st.dataframe(flip_display, use_container_width=True)

                row = flip_row.iloc[0]
                comps_matched = match_comps(row, df_comps_focus)
                st.markdown("#### Matching Comps:")
                if not comps_matched.empty:
                    sell_date_col = None
                    for candidate in ["Close Dt", "Close Date", "Sold Date"]:
                        if candidate in comps_matched.columns:
                            sell_date_col = candidate
                            break
                    if not sell_date_col:
                        for col in comps_matched.columns:
                            if "date" in col.lower():
                                sell_date_col = col
                                break

                    comps_cols = [
                        "MLS #", "Status", "Address", "County", "City", "Zip", "Sub",
                        "Bedrooms", "Total Finished SF", "Sale Price"
                    ]
                    if "Property Tax" in comps_matched.columns:
                        sp_idx = comps_cols.index("Sale Price")
                        comps_cols.insert(sp_idx + 1, "Property Tax")
                    if sell_date_col:
                        comps_cols.append(sell_date_col)
                    if "Zillow" not in comps_cols:
                        comps_cols.append("Zillow")
                    existing_comp_cols = [col for col in comps_cols if col in comps_matched.columns]

                    comps_display = comps_matched[existing_comp_cols].copy()
                    if "Sale Price" in comps_display.columns:
                        comps_display["Sale Price"] = comps_display["Sale Price"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                    if "Property Tax" in comps_display.columns:
                        comps_display["Property Tax"] = comps_display["Property Tax"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                    if "Total Finished SF" in comps_display.columns:
                        comps_display["Total Finished SF"] = comps_display["Total Finished SF"].map(lambda x: "{:,.0f}".format(x) if pd.notnull(x) else x)
                    if sell_date_col and sell_date_col in comps_display.columns and sell_date_col != "Close Dt":
                        comps_display.rename(columns={sell_date_col: "Close Dt"}, inplace=True)
                        existing_comp_cols = [
                            col if col != sell_date_col else "Close Dt"
                            for col in existing_comp_cols
                        ]

                    def make_zillow_btn(address):
                        if pd.isnull(address) or address.strip() == "":
                            return ""
                        url = f"https://www.zillow.com/homes/{address.replace(' ', '-').replace(',', '')}_rb/"
                        return f"<a href='{url}' target='_blank'><button style='background:#2563eb;color:white;padding:2px 8px;border-radius:5px;border:none;'>Zillow</button></a>"

                    comps_display["Zillow"] = comps_display["Address"].apply(make_zillow_btn)
                    if "Zillow" in comps_display.columns and "Zillow" not in existing_comp_cols:
                        existing_comp_cols.append("Zillow")
                    st.write(comps_display[existing_comp_cols].to_html(escape=False, index=False), unsafe_allow_html=True)

                    # ---- RAW EXPORT ALL 4 TABS ----
                    criteria_dict = {
                        "Same ZIP": match_zip,
                        "Same County": match_county,
                        "Same City": match_city,
                        "Same Sub": match_sub,
                        "Same # Bedrooms": match_beds,
                        "SF Range (%)": sf_pct
                    }
                    # ADD ZILLOW HYPERLINK TO FOCUSED FLIPS
                    df_props_export = df_props.copy()
                    if "Address" in df_props_export.columns:
                        df_props_export["Zillow"] = df_props_export["Address"].apply(build_zillow_hyperlink)
                    # ADD ZILLOW HYPERLINK TO FLIP ROW FOR FLIP+COMPS EXPORT
                    raw_flip = flip_row[existing_flip_cols].copy()
                    if "Address" in raw_flip.columns:
                        raw_flip["Zillow"] = raw_flip["Address"].apply(build_zillow_hyperlink)
                    if "Zillow" not in existing_flip_cols:
                        existing_flip_cols.append("Zillow")
                    # Only include columns that actually exist in comps_matched!
                    export_cols = [c for c in existing_comp_cols if c != "Zillow" and c in comps_matched.columns]
                    raw_comps = comps_matched[export_cols].copy()
                    # Also add Zillow hyperlinks to comps if desired (optional)
                    if "Address" in raw_comps.columns:
                        raw_comps["Zillow"] = raw_comps["Address"].apply(build_zillow_hyperlink)
                    excel_export = export_full_analysis(
                        summary_table,
                        df_props_export,
                        raw_flip,
                        raw_comps,
                        criteria_dict
                    )
                    st.download_button(
                        label="⬇️ Export All (Excel, RAW, 4 Tabs)",
                        data=excel_export,
                        file_name=f"MLS_{mls_num}_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("No comps found for this property with the current criteria.")
else:
    st.info("Upload both Listings and Comps CSV files to get started.")
