import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Monthly Production Aggregator", page_icon="🏭", layout="centered")

st.title("🏭 Monthly Production Aggregator")
st.write("Drag and drop your monthly production file below. The app will automatically scan all sheets, sum the production quantities per SKU, and prepare an Accounting-ready Excel file.")

# 1. Department Selector (Now includes Hot Kitchen)
department = st.radio("Which department are you processing?", ("RTE / RTD", "Bakery", "Hot Kitchen & Tortilla"))

uploaded_file = st.file_uploader(f"Drop your {department} Excel file here", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info("File uploaded successfully! Processing daily sheets...")
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        progress_bar = st.progress(0)
        
        # ==========================================
        # ROUTE A: BAKERY & RTE (Standard Single Table)
        # ==========================================
        if department in ["RTE / RTD", "Bakery"]:
            all_sheets_data = []
            total_quantity_found = 0
            total_rows_found = 0
            
            for i, sheet in enumerate(sheet_names):
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                
                header_row_idx = -1
                for idx, row in df_raw.iterrows():
                    row_str = row.astype(str).str.lower()
                    if row_str.str.contains('item description').any() and (row_str.str.contains('qty').any() or row_str.str.contains('quantity').any()):
                        header_row_idx = idx
                        break
                
                if header_row_idx != -1:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)
                else:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet)

                item_col = next((col for col in df_sheet.columns if 'item description' in str(col).lower()), None)
                qty_col = next((col for col in df_sheet.columns if str(col).lower().strip() in ['qty', 'quantity']), None)

                if item_col and qty_col:
                    df_extracted = df_sheet[[item_col, qty_col]].copy()
                    df_extracted.rename(columns={item_col: 'Item Description', qty_col: 'Qty'}, inplace=True)
                    
                    df_extracted['Source Sheet (Date)'] = sheet
                    df_extracted['Qty'] = pd.to_numeric(df_extracted['Qty'], errors='coerce')
                    df_extracted = df_extracted.dropna(subset=['Qty'])  
                    df_extracted = df_extracted[df_extracted['Qty'] > 0] 
                    df_extracted['Item Description'] = df_extracted['Item Description'].astype(str).str.strip()
                    
                    all_sheets_data.append(df_extracted)
                    
                    total_quantity_found += df_extracted['Qty'].sum()
                    total_rows_found += len(df_extracted)
                
                progress_bar.progress((i + 1) / len(sheet_names))

            if not all_sheets_data:
                st.error("❌ Could not find valid production data.")
            else:
                df_master = pd.concat(all_sheets_data, ignore_index=True)
                monthly_totals = df_master.groupby('Item Description')['Qty'].sum().reset_index().sort_values(by='Item Description')
                aggregated_sum = monthly_totals['Qty'].sum()
                df_audit = df_master[['Source Sheet (Date)', 'Item Description', 'Qty']]
                
                st.success("✅ Parsing Complete!")
                
                st.subheader("📊 Accounting Validation")
                col1, col2, col3 = st.columns(3)
                col1.metric("Unique SKUs Produced", len(monthly_totals))
                col2.metric("Total Line Items", total_rows_found)
                col3.metric("Total Qty Produced", aggregated_sum)
                
                if total_quantity_found == aggregated_sum:
                    st.caption("✅ Validation Passed: Raw sum exactly matches aggregated sum.")
                else:
                    st.error("❌ Validation Failed: Discrepancy detected.")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    monthly_totals.to_excel(writer, index=False, sheet_name="Monthly Summary")
                    df_audit.to_excel(writer, index=False, sheet_name="Audit Log")
                
                processed_data = output.getvalue()
                safe_dept = department.replace(" / ", "_").replace(" ", "_")
                
                st.markdown("---")
                st.download_button(
                    label=f"📥 Download Total_Monthly_{safe_dept}_Production.xlsx",
                    data=processed_data,
                    file_name=f"Total_Monthly_{safe_dept}_Production.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ==========================================
        # ROUTE B: HOT KITCHEN & TORTILLA (Stacked Tables)
        # ==========================================
        elif department == "Hot Kitchen & Tortilla":
            hk_data = []
            tortilla_data = []
            
            for i, sheet in enumerate(sheet_names):
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                
                hk_start = -1
                tortilla_start = -1
                
                # Find where each table starts
                for idx, row in df_raw.iterrows():
                    row_str = row.astype(str).str.lower().str.strip()
                    if row_str.str.contains('hot kitchen production').any():
                        hk_start = idx
                    elif row_str.str.contains('tortilla production').any():
                        tortilla_start = idx

                def clean_section(df_section, section_name):
                    if df_section.empty or len(df_section) < 2:
                        return None
                    # Set the first row as the header
                    df_section.columns = df_section.iloc[0].astype(str).str.strip().str.title()
                    df_section = df_section[1:].copy()
                    
                    # Ensure required columns exist
                    if not all(col in df_section.columns for col in ['Item', 'Qty', 'Uom']):
                        # Fallback for UOM capitalization
                        if 'UOM' in df_section.columns:
                            df_section.rename(columns={'UOM': 'Uom'}, inplace=True)
                        elif 'U/M' in df_section.columns:
                            df_section.rename(columns={'U/M': 'Uom'}, inplace=True)
                            
                    if 'Item' in df_section.columns and 'Qty' in df_section.columns:
                        df_section['Source Sheet (Date)'] = sheet
                        df_section = df_section.dropna(subset=['Item', 'Qty'])
                        
                        # Remove "Total" rows
                        df_section = df_section[~df_section['Item'].astype(str).str.lower().str.contains('total')]
                        
                        df_section['Qty'] = pd.to_numeric(df_section['Qty'], errors='coerce')
                        df_section = df_section.dropna(subset=['Qty'])
                        df_section = df_section[df_section['Qty'] > 0]
                        df_section['Item'] = df_section['Item'].astype(str).str.strip()
                        
                        # Ensure UOM exists for grouping
                        if 'Uom' not in df_section.columns:
                            df_section['Uom'] = 'N/A'
                            
                        return df_section[['Source Sheet (Date)', 'Item', 'Qty', 'Uom']]
                    return None

                # Slice and process Hot Kitchen
                if hk_start != -1:
                    hk_end = tortilla_start if tortilla_start != -1 else len(df_raw)
                    hk_slice = df_raw.iloc[hk_start+1:hk_end].copy()
                    cleaned_hk = clean_section(hk_slice, "Hot Kitchen")
                    if cleaned_hk is not None:
                        hk_data.append(cleaned_hk)

                # Slice and process Tortilla
                if tortilla_start != -1:
                    tortilla_slice = df_raw.iloc[tortilla_start+1:].copy()
                    cleaned_tortilla = clean_section(tortilla_slice, "Tortilla")
                    if cleaned_tortilla is not None:
                        tortilla_data.append(cleaned_tortilla)

                progress_bar.progress((i + 1) / len(sheet_names))

            if not hk_data and not tortilla_data:
                st.error("❌ Could not find valid Hot Kitchen or Tortilla data.")
            else:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    
                    if hk_data:
                        df_hk_master = pd.concat(hk_data, ignore_index=True)
                        hk_summary = df_hk_master.groupby(['Item', 'Uom'])['Qty'].sum().reset_index().sort_values(by='Item')
                        hk_summary.to_excel(writer, index=False, sheet_name="Hot Kitchen Summary")
                        df_hk_master.to_excel(writer, index=False, sheet_name="Hot Kitchen Audit Log")
                        st.success(f"✅ Hot Kitchen parsed: {len(hk_summary)} SKUs found.")
                        
                    if tortilla_data:
                        df_tortilla_master = pd.concat(tortilla_data, ignore_index=True)
                        tortilla_summary = df_tortilla_master.groupby(['Item', 'Uom'])['Qty'].sum().reset_index().sort_values(by='Item')
                        tortilla_summary.to_excel(writer, index=False, sheet_name="Tortilla Summary")
                        df_tortilla_master.to_excel(writer, index=False, sheet_name="Tortilla Audit Log")
                        st.success(f"✅ Tortilla parsed: {len(tortilla_summary)} SKUs found.")

                processed_data = output.getvalue()
                
                st.markdown("---")
                st.download_button(
                    label="📥 Download Total_Monthly_HK_and_Tortilla.xlsx",
                    data=processed_data,
                    file_name="Total_Monthly_HK_and_Tortilla_Production.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")