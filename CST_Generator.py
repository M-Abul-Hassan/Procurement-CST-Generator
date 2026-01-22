"""
Streamlit-based Comparative Statement (CST) Generator for Procurement
Author: AI Assistant
Version: 1.0
"""

import streamlit as st
import pandas as pd
import pdfplumber
import openpyxl
from io import BytesIO
import re

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_from_pdf(file):
    """
    Extract quotation data from PDF file
    Returns: List of dictionaries with product information
    """
    try:
        items = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # Extract table data
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        # Skip header row
                        for row in table[1:]:
                            if len(row) >= 5 and row[0]:  # Ensure row has data
                                try:
                                    items.append({
                                        'Product ID': str(row[0]).strip(),
                                        'Product Name': str(row[1]).strip(),
                                        'Drawing Number': str(row[2]).strip(),
                                        'Quantity': float(row[3]),
                                        'Unit Price': float(row[4])
                                    })
                                except (ValueError, IndexError):
                                    continue
        return items
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return []


def extract_from_excel(file):
    """
    Extract quotation data from Excel file
    Returns: List of dictionaries with product information
    """
    try:
        df = pd.read_excel(file)
        
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Create a mapping for column name variations (case-insensitive)
        col_mapping = {
            'Product ID': ['product id', 'product_id', 'prod_id', 'item_id', 'id', 'productid'],
            'Product Name': ['product name', 'product_name', 'prod_name', 'item_name', 'name', 'description', 'productname'],
            'Drawing Number': ['drawing number', 'drawing_number', 'drawing_no', 'drg_no', 'drawing', 'drawingnumber', 'drg number'],
            'Quantity': ['quantity', 'qty', 'amount', 'qnty'],
            'Unit Price': ['unit price', 'unit_price', 'unit_cost', 'price', 'rate', 'unitprice', 'unit cost']
        }
        
        # Find the actual column names in the dataframe
        rename_dict = {}
        for standard_name, variations in col_mapping.items():
            for col in df.columns:
                # Check exact match first (case-insensitive)
                if col.lower() in variations:
                    rename_dict[col] = standard_name
                    break
                # Check if any variation is contained in the column name
                elif any(v in col.lower() for v in variations):
                    rename_dict[col] = standard_name
                    break
        
        # Rename columns to standard names
        df = df.rename(columns=rename_dict)
        
        # Check if all required columns are present
        required_cols = ['Product ID', 'Product Name', 'Drawing Number', 'Quantity', 'Unit Price']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"Missing columns in Excel file: {', '.join(missing_cols)}")
            st.info(f"Available columns: {', '.join(df.columns.tolist())}")
            return []
        
        # Select required columns and drop rows with missing Product ID
        df = df[required_cols].dropna(subset=['Product ID'])
        
        # Clean and convert data types
        df['Product ID'] = df['Product ID'].astype(str).str.strip()
        df['Product Name'] = df['Product Name'].astype(str).str.strip()
        df['Drawing Number'] = df['Drawing Number'].astype(str).str.strip()
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
        df['Unit Price'] = pd.to_numeric(df['Unit Price'], errors='coerce')
        
        # Remove rows with invalid numeric values
        df = df.dropna(subset=['Quantity', 'Unit Price'])
        
        # Convert to list of dictionaries
        items = df.to_dict('records')
        
        return items
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return []


def normalize_data(suppliers_data):
    """
    Normalize and merge all supplier quotations
    Args:
        suppliers_data: List of tuples (supplier_name, quality_rating, items)
    Returns:
        DataFrame with normalized CST data
    """
    # Collect all unique products across all suppliers
    all_products = {}
    
    for supplier_name, quality_rating, items in suppliers_data:
        for item in items:
            # Create composite key
            key = (
                str(item['Product ID']).strip(),
                str(item['Product Name']).strip(),
                str(item['Drawing Number']).strip()
            )
            
            if key not in all_products:
                all_products[key] = {
                    'Product ID': item['Product ID'],
                    'Product Name': item['Product Name'],
                    'Drawing Number': item['Drawing Number'],
                    'Quantity': item['Quantity']
                }
    
    # Create base DataFrame
    df = pd.DataFrame(list(all_products.values()))
    
    # Add supplier columns
    for supplier_name, quality_rating, items in suppliers_data:
        # Create a mapping for this supplier
        supplier_prices = {}
        for item in items:
            key = (
                str(item['Product ID']).strip(),
                str(item['Product Name']).strip(),
                str(item['Drawing Number']).strip()
            )
            supplier_prices[key] = item['Unit Price']
        
        # Add unit price column
        col_name = f"{supplier_name} (Q{quality_rating})"
        df[col_name] = df.apply(
            lambda row: supplier_prices.get(
                (row['Product ID'], row['Product Name'], row['Drawing Number']),
                0.0
            ),
            axis=1
        )
        
        # Add total price column
        total_col_name = f"{supplier_name} Total"
        df[total_col_name] = df[col_name] * df['Quantity']
    
    return df


def apply_conditional_formatting(df, supplier_cols, highlight_lowest, highlight_quality, 
                                 manual_selections, suppliers_data):
    """
    Apply conditional formatting to dataframe for display
    Returns: Styled DataFrame
    """
    def highlight_cells(row):
        styles = [''] * len(row)
        
        # Get supplier price columns (excluding total columns)
        price_cols = [col for col in supplier_cols if 'Total' not in col]
        prices = row[price_cols]
        
        # Find lowest price (excluding zeros)
        non_zero_prices = prices[prices > 0]
        if len(non_zero_prices) > 0:
            min_price = non_zero_prices.min()
        else:
            min_price = None
        
        # Find best quality supplier
        best_quality_supplier = None
        max_quality = 0
        for supplier_name, quality_rating, _ in suppliers_data:
            if quality_rating > max_quality:
                max_quality = quality_rating
                best_quality_supplier = f"{supplier_name} (Q{quality_rating})"
        
        # Apply styling
        for idx, col in enumerate(row.index):
            if col in price_cols:
                supplier_col = col
                
                # Check manual selection
                if row['Product ID'] in manual_selections:
                    if manual_selections[row['Product ID']] in col:
                        styles[idx] = 'background-color: #ADD8E6'  # Blue
                        continue
                
                # Highlight lowest price
                if highlight_lowest and min_price and row[col] == min_price and row[col] > 0:
                    styles[idx] = 'background-color: #90EE90'  # Green
                
                # Highlight best quality
                elif highlight_quality and supplier_col == best_quality_supplier and row[col] > 0:
                    styles[idx] = 'background-color: #FFFFE0'  # Yellow
        
        return styles
    
    return df.style.apply(highlight_cells, axis=1)


def export_to_excel(df):
    """
    Export DataFrame to Excel with formatting
    Returns: BytesIO object
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comparative Statement')
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Comparative Statement']
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    output.seek(0)
    return output


# ============================================================================
# STREAMLIT APP
# ============================================================================

def main():
    st.set_page_config(
        page_title="Procurement CST Generator",
        page_icon="📊",
        layout="wide"
    )
    
    # Custom CSS
    st.markdown("""
        <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1f77b4;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1.2rem;
            color: #666;
            margin-bottom: 2rem;
        }
        .supplier-card {
            background-color: #f0f2f6;
            padding: 1rem;
            border-radius: 0.5rem;
            margin: 0.5rem 0;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.markdown('<div class="main-header">📊 Procurement CST Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Upload supplier quotations and generate comparative statements</div>', 
                unsafe_allow_html=True)
    
    # Initialize session state
    if 'suppliers_data' not in st.session_state:
        st.session_state.suppliers_data = []
    if 'cst_df' not in st.session_state:
        st.session_state.cst_df = None
    if 'manual_selections' not in st.session_state:
        st.session_state.manual_selections = {}
    
    # Sidebar - Upload Section
    with st.sidebar:
        st.header("📤 Upload Quotations")
        
        # Upload mode selector
        upload_mode = st.radio(
            "Upload Mode",
            options=["Single File", "Multiple Files"],
            horizontal=True
        )
        
        if upload_mode == "Single File":
            # Single file upload mode
            supplier_name = st.text_input("Supplier Name", key="supplier_name")
            quality_rating = st.selectbox(
                "Quality Rating",
                options=[5, 4, 3, 2, 1],
                format_func=lambda x: f"{x} - {'Excellent' if x==5 else 'Very Good' if x==4 else 'Good' if x==3 else 'Fair' if x==2 else 'Poor'}",
                key="quality_rating"
            )
            
            uploaded_file = st.file_uploader(
                "Upload Quotation File",
                type=['pdf', 'xlsx', 'xls'],
                key="file_uploader"
            )
            
            if st.button("Add Supplier", type="primary"):
                if not supplier_name:
                    st.error("Please enter supplier name")
                elif not uploaded_file:
                    st.error("Please upload a file")
                else:
                    # Extract data based on file type
                    if uploaded_file.name.endswith('.pdf'):
                        items = extract_from_pdf(uploaded_file)
                    else:
                        items = extract_from_excel(uploaded_file)
                    
                    if items:
                        st.session_state.suppliers_data.append(
                            (supplier_name, quality_rating, items)
                        )
                        st.success(f"✅ Added {supplier_name} with {len(items)} items")
                        st.rerun()
                    else:
                        st.error("No data extracted from file. Please check file format.")
        
        else:
            # Multiple files upload mode
            st.info("💡 Upload multiple files at once. You'll configure each supplier after upload.")
            
            uploaded_files = st.file_uploader(
                "Upload Multiple Quotation Files",
                type=['pdf', 'xlsx', 'xls'],
                accept_multiple_files=True,
                key="multi_file_uploader"
            )
            
            if uploaded_files:
                st.markdown(f"**{len(uploaded_files)} file(s) uploaded**")
                
                # Initialize session state for file configurations
                if 'file_configs' not in st.session_state:
                    st.session_state.file_configs = {}
                
                # Create configuration for each file
                for file in uploaded_files:
                    file_key = file.name
                    
                    with st.expander(f"📄 {file.name}", expanded=True):
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            supplier_name = st.text_input(
                                "Supplier Name",
                                key=f"name_{file_key}",
                                placeholder="e.g., ABC Corp"
                            )
                        
                        with col2:
                            quality_rating = st.selectbox(
                                "Quality",
                                options=[5, 4, 3, 2, 1],
                                format_func=lambda x: f"{x}⭐",
                                key=f"rating_{file_key}"
                            )
                        
                        # Store configuration
                        st.session_state.file_configs[file_key] = {
                            'name': supplier_name,
                            'rating': quality_rating,
                            'file': file
                        }
                
                st.divider()
                
                if st.button("📥 Process All Files", type="primary"):
                    success_count = 0
                    error_count = 0
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for idx, (file_key, config) in enumerate(st.session_state.file_configs.items()):
                        if not config['name']:
                            st.warning(f"⚠️ Skipped {file_key}: No supplier name provided")
                            error_count += 1
                            continue
                        
                        status_text.text(f"Processing {config['name']}...")
                        
                        # Extract data based on file type
                        file = config['file']
                        if file.name.endswith('.pdf'):
                            items = extract_from_pdf(file)
                        else:
                            items = extract_from_excel(file)
                        
                        if items:
                            st.session_state.suppliers_data.append(
                                (config['name'], config['rating'], items)
                            )
                            success_count += 1
                        else:
                            st.error(f"❌ {file_key}: No data extracted")
                            error_count += 1
                        
                        progress_bar.progress((idx + 1) / len(st.session_state.file_configs))
                    
                    status_text.empty()
                    progress_bar.empty()
                    
                    if success_count > 0:
                        st.success(f"✅ Successfully added {success_count} supplier(s)!")
                        # Clear file configs
                        st.session_state.file_configs = {}
                        st.rerun()
                    
                    if error_count > 0:
                        st.warning(f"⚠️ {error_count} file(s) had errors")
        
        st.divider()
        
        # Display uploaded suppliers
        if st.session_state.suppliers_data:
            st.subheader("📋 Uploaded Suppliers")
            for idx, (name, rating, items) in enumerate(st.session_state.suppliers_data):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(f"**{name}** (Quality: {rating})")
                    st.caption(f"{len(items)} items")
                with col2:
                    if st.button("🗑️", key=f"delete_{idx}"):
                        st.session_state.suppliers_data.pop(idx)
                        st.session_state.cst_df = None
                        st.rerun()
            
            st.divider()
            
            if st.button("🔄 Clear All", type="secondary"):
                st.session_state.suppliers_data = []
                st.session_state.cst_df = None
                st.session_state.manual_selections = {}
                st.rerun()
    
    # Main Area
    if st.session_state.suppliers_data:
        # Generate CST button
        if st.button("🎯 Generate Comparative Statement", type="primary"):
            with st.spinner("Generating CST..."):
                st.session_state.cst_df = normalize_data(st.session_state.suppliers_data)
                st.success("✅ CST Generated Successfully!")
        
        # Display CST
        if st.session_state.cst_df is not None:
            df = st.session_state.cst_df
            
            # Controls
            st.subheader("⚙️ Display Options")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                highlight_lowest = st.checkbox("🟢 Highlight Lowest Price", value=True)
            with col2:
                highlight_quality = st.checkbox("🟡 Highlight Best Quality", value=True)
            with col3:
                st.markdown("🔵 Manual Selection (via dropdown)")
            
            st.divider()
            
            # Manual selection dropdowns
            st.subheader("✋ Manual Supplier Selection")
            supplier_names = [name for name, _, _ in st.session_state.suppliers_data]
            
            cols = st.columns(4)
            for idx, row in df.iterrows():
                col_idx = idx % 4
                with cols[col_idx]:
                    product_id = row['Product ID']
                    selected = st.selectbox(
                        f"{product_id}",
                        options=["Auto"] + supplier_names,
                        key=f"select_{product_id}_{idx}"
                    )
                    if selected != "Auto":
                        st.session_state.manual_selections[product_id] = selected
                    elif product_id in st.session_state.manual_selections:
                        del st.session_state.manual_selections[product_id]
            
            st.divider()
            
            # Display CST Table
            st.subheader("📊 Comparative Statement")
            
            # Get supplier columns
            supplier_cols = [col for col in df.columns 
                           if col not in ['Product ID', 'Product Name', 'Drawing Number', 'Quantity']]
            
            # Apply styling
            styled_df = apply_conditional_formatting(
                df, supplier_cols, highlight_lowest, highlight_quality,
                st.session_state.manual_selections, st.session_state.suppliers_data
            )
            
            st.dataframe(styled_df, use_container_width=True, height=600)
            
            # Export Section
            st.divider()
            st.subheader("💾 Export Options")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Export to Excel
                excel_data = export_to_excel(df)
                st.download_button(
                    label="📥 Download as Excel",
                    data=excel_data,
                    file_name="comparative_statement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            
            with col2:
                # Export to CSV
                csv_data = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv_data,
                    file_name="comparative_statement.csv",
                    mime="text/csv",
                    type="secondary"
                )
    
    else:
        # Empty state
        st.info("👈 Upload supplier quotations from the sidebar to get started")
        
        st.markdown("### 📝 Instructions")
        st.markdown("""
        1. **Upload Files**: Use the sidebar to upload PDF or Excel quotation files
        2. **Enter Details**: Provide supplier name and quality rating for each upload
        3. **Generate CST**: Click the generate button to create the comparative statement
        4. **Review & Select**: Use highlighting and manual selection to choose suppliers
        5. **Export**: Download the final CST in Excel or CSV format
        
        **Supported Formats:**
        - PDF files with tabular data
        - Excel files (.xlsx, .xls) with columns: Product ID, Product Name, Drawing Number, Quantity, Unit Price
        """)


if __name__ == "__main__":
    main()