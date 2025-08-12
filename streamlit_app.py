#!/usr/bin/env python3
"""
Streamlit Web Interface for R2D Reconciliation Tool
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime
import traceback

# Import the r2d_recon module
try:
    import r2d_recon
except ImportError as e:
    st.error(f"Could not import r2d_recon module: {e}")
    st.stop()

def main():
    st.title("🔍 R2D Reconciliation Tool")
    st.markdown("Upload your Excel files and get reconciliation results instantly!")
    
    # File upload section
    st.header("📂 Upload Files")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file containing both R2D and Chase sheets",
        type=['xlsx', 'xls'],
        help="Upload an Excel file with multiple sheets"
    )
    
    if uploaded_file is None:
        st.info("👆 Please upload an Excel file to get started")
        return
    
    # Show file info
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    # Configuration section
    st.header("⚙️ Configuration")
    
    col1, col2 = st.columns(2)
    
    with col1:
        r2d_sheet = st.text_input(
            "R2D Sheet Name",
            value="Repayments to Date",
            help="Name of the sheet containing R2D data"
        )
    
    with col2:
        chase_sheet = st.text_input(
            "Chase Sheet Name", 
            value="Chase",
            help="Name of the sheet containing Chase data"
        )
    
    # Optional date filter
    st.subheader("📅 Optional Date Filter")
    use_date_filter = st.checkbox("Ignore debits before a specific date")
    ignore_debits_before = None
    
    if use_date_filter:
        ignore_debits_before = st.date_input(
            "Ignore debits before this date",
            help="Only transactions on or after this date will be processed"
        )
        if ignore_debits_before:
            ignore_debits_before = ignore_debits_before.strftime("%Y-%m-%d")
    
    # Process button
    st.header("🚀 Run Reconciliation")
    
    if st.button("Run Reconciliation", type="primary", use_container_width=True):
        run_reconciliation(uploaded_file, r2d_sheet, chase_sheet, ignore_debits_before)

def run_reconciliation(uploaded_file, r2d_sheet, chase_sheet, ignore_debits_before):
    """Run the reconciliation process with progress tracking"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Save uploaded file to temp location
        status_text.text("📁 Processing uploaded file...")
        progress_bar.progress(10)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
            tmp_input.write(uploaded_file.getvalue())
            input_path = tmp_input.name
        
        # Step 2: Create output file path
        status_text.text("📝 Preparing output file...")
        progress_bar.progress(20)
        
        timestamp = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"Repayments_to_Date_recon-{timestamp}.xlsx"
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
            output_path = tmp_output.name
        
        # Step 3: Run reconciliation
        status_text.text("🔄 Running reconciliation algorithm...")
        progress_bar.progress(40)
        
        # Call the r2d_recon.run function
        r2d_recon.run(
            file_path=input_path,
            r2d_sheet=r2d_sheet,
            chase_sheet=chase_sheet, 
            out_path=output_path,
            ignore_debits_before=ignore_debits_before
        )
        
        progress_bar.progress(80)
        status_text.text("✅ Reconciliation complete! Preparing download...")
        
        # Step 4: Provide download
        progress_bar.progress(100)
        
        # Read the output file for download
        with open(output_path, 'rb') as f:
            output_data = f.read()
        
        status_text.text("🎉 Success! Download your reconciliation results below.")
        
        st.download_button(
            label="📥 Download Reconciliation Results",
            data=output_data,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Show success message
        st.success("✅ Reconciliation completed successfully!")
        st.balloons()
        
        # Display some summary info if possible
        try:
            # Try to read and display some summary info
            df_summary = pd.read_excel(output_path, sheet_name=0, nrows=5)
            st.subheader("📊 Preview of Results")
            st.dataframe(df_summary)
        except Exception:
            # If we can't read the summary, that's okay
            pass
            
    except Exception as e:
        st.error(f"❌ An error occurred during reconciliation:")
        st.code(str(e))
        
        # Show detailed error for debugging
        with st.expander("🔍 Detailed Error Information"):
            st.code(traceback.format_exc())
    
    finally:
        # Clean up temporary files
        try:
            if 'input_path' in locals():
                os.unlink(input_path)
            if 'output_path' in locals():
                os.unlink(output_path)
        except Exception:
            pass

def show_sidebar_info():
    """Show helpful information in the sidebar"""
    st.sidebar.header("ℹ️ How to Use")
    st.sidebar.markdown("""
    1. **Upload** your Excel file containing both R2D and Chase data
    2. **Configure** the sheet names (defaults usually work)
    3. **Optionally** set a date filter for debits
    4. **Click** "Run Reconciliation"
    5. **Download** your results when complete
    """)
    
    st.sidebar.header("📋 Requirements")
    st.sidebar.markdown("""
    - Excel file (.xlsx or .xls)
    - Sheet with R2D data (default: "Repayments to Date")
    - Sheet with Chase data (default: "Chase")
    """)
    
    st.sidebar.header("🆘 Need Help?")
    st.sidebar.markdown("""
    - Make sure your Excel file has the correct sheet names
    - Check that your data has the expected column headers
    - Contact support if you encounter errors
    """)

if __name__ == "__main__":
    # Configure page
    st.set_page_config(
        page_title="R2D Reconciliation Tool",
        page_icon="🔍",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Show sidebar info
    show_sidebar_info()
    
    # Run main app
    main()
