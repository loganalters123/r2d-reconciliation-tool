#!/usr/bin/env python3
"""
ClaimAngel 1160 Reconciliation Tool - Modern Streamlit Interface
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from datetime import datetime
import traceback
from pathlib import Path

# Import the r2d_recon module
try:
    import r2d_recon
except ImportError as e:
    st.error(f"Could not import r2d_recon module: {e}")
    st.stop()

def inject_custom_css():
    """Inject custom CSS for modern branding and styling"""
    st.markdown("""
         # Download button
        st.download_button(
            label="📥 Download Reconciliation Results",
            data=output_data,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Balance Reconciliation Section
        st.markdown('</div>', unsafe_allow_html=True)
        show_balance_reconciliation(output_path)
        st.markdown('<div class="ca-card">', unsafe_allow_html=True)
        
        # Show balloons animation
        st.balloons()    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300..900&display=swap');
    
    :root {
        --ink: #101721;
        --teal-dark: #123C40;
        --accent: #178CC4;
        --accent-light: #68BCE4;
        --sky: #BEDFEE;
        --sand: #E6E4E1;
        --cloud: #F1F1F1;
    }
    
    html, body, [class*="appview-container"] {
        font-family: 'Outfit', system-ui, -apple-system, Segoe UI, Roboto, 'Helvetica Neue', Arial, sans-serif !important;
        color: var(--ink) !important;
    }
    
    .main .block-container {
        padding-top: 2rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 1200px !important;
    }
    
    .ca-header {
        display: flex;
        align-items: center;
        margin-bottom: 2rem;
        padding: 1.5rem;
        background: white;
        border: 1px solid var(--sand);
        border-radius: 20px;
        box-shadow: 0 8px 25px rgba(16,23,33,0.08);
    }
    
    .ca-header img {
        height: 52px;
        margin-right: 1.5rem;
    }
    
    .ca-header-text h1 {
        margin: 0 0 0.25rem 0 !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.3px !important;
        color: var(--ink) !important;
    }
    
    .ca-header-text p {
        margin: 0 !important;
        color: #445566 !important;
        opacity: 0.9 !important;
        font-size: 1.1rem !important;
        font-weight: 400 !important;
    }
    
    .ca-card {
        background: white !important;
        border: 1px solid var(--sand) !important;
        border-radius: 20px !important;
        box-shadow: 0 8px 25px rgba(16,23,33,0.06) !important;
        padding: 2rem !important;
        margin: 1.5rem 0 !important;
    }
    
    .ca-file-upload {
        border: 3px dashed var(--sand) !important;
        border-radius: 16px !important;
        padding: 2rem !important;
        text-align: center !important;
        transition: all 0.3s ease !important;
        background: #fafbfc !important;
    }
    
    .ca-file-upload:hover {
        border-color: var(--accent-light) !important;
        background: white !important;
        transform: translateY(-2px) !important;
    }
    
    /* Streamlit button styling */
    div.stButton > button {
        border-radius: 14px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        border: 2px solid transparent !important;
        background: var(--accent) !important;
        color: white !important;
        transition: all 0.3s ease !important;
        font-family: 'Outfit', sans-serif !important;
        letter-spacing: 0.3px !important;
    }
    
    div.stButton > button:hover {
        background: var(--teal-dark) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(23,140,196,0.3) !important;
    }
    
    .stDownloadButton button {
        border-radius: 14px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        border: 2px solid var(--accent) !important;
        background: var(--accent) !important;
        color: white !important;
        transition: all 0.3s ease !important;
        font-family: 'Outfit', sans-serif !important;
        letter-spacing: 0.3px !important;
    }
    
    .stDownloadButton button:hover {
        background: var(--teal-dark) !important;
        border-color: var(--teal-dark) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(23,140,196,0.3) !important;
    }
    
    /* Input styling */
    .stTextInput input, .stSelectbox select {
        border-radius: 12px !important;
        border: 2px solid var(--sand) !important;
        padding: 0.7rem !important;
        font-family: 'Outfit', sans-serif !important;
        font-size: 1rem !important;
    }
    
    .stTextInput input:focus, .stSelectbox select:focus {
        border-color: var(--accent-light) !important;
        box-shadow: 0 0 0 3px rgba(23,140,196,0.1) !important;
    }
    
    /* Status pills */
    .pill {
        display: inline-block !important;
        padding: 0.4rem 1rem !important;
        border-radius: 25px !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
        margin: 0.5rem 0.5rem 0.5rem 0 !important;
    }
    
    .pill-success {
        background: var(--sky) !important;
        color: var(--teal-dark) !important;
        border: 1px solid var(--accent-light) !important;
    }
    
    .pill-info {
        background: #e8f4fd !important;
        color: var(--accent) !important;
        border: 1px solid var(--accent-light) !important;
    }
    
    .pill-warning {
        background: #fff3cd !important;
        color: #856404 !important;
        border: 1px solid #ffeaa7 !important;
    }
    
    /* File uploader specific */
    .stFileUploader {
        border: none !important;
    }
    
    .stFileUploader > div {
        border: 3px dashed var(--sand) !important;
        border-radius: 16px !important;
        padding: 2rem !important;
        background: #fafbfc !important;
        transition: all 0.3s ease !important;
    }
    
    .stFileUploader > div:hover {
        border-color: var(--accent-light) !important;
        background: white !important;
    }
    
    /* Progress bar */
    .stProgress .st-bo {
        background-color: var(--sky) !important;
    }
    
    .stProgress .st-bp {
        background-color: var(--accent) !important;
    }
    
    /* Success/error messages */
    .stSuccess {
        background-color: var(--sky) !important;
        border: 1px solid var(--accent-light) !important;
        border-radius: 12px !important;
    }
    
    .stError {
        border-radius: 12px !important;
    }
    
    /* Section headers */
    h1, h2, h3 {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
        color: var(--ink) !important;
        letter-spacing: 0.2px !important;
    }
    
    h2 {
        font-size: 1.6rem !important;
        margin-bottom: 1rem !important;
        margin-top: 2rem !important;
    }
    
    h3 {
        font-size: 1.3rem !important;
        margin-bottom: 0.8rem !important;
        margin-top: 1.5rem !important;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .ca-logo-fallback {
        display: inline-block;
        background: var(--accent);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-weight: 700;
        font-size: 1.2rem;
        letter-spacing: 0.5px;
        margin-right: 1.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

def show_header():
    """Display the branded header"""
    logo_path = Path(__file__).parent / "assets" / "claimangel_logo.png"
    
    if logo_path.exists():
        # Use actual logo
        logo_html = f'<img src="data:image/png;base64,{get_logo_base64()}" alt="ClaimAngel Logo">'
    else:
        # Fallback text logo
        logo_html = '<div class="ca-logo-fallback">CA</div>'
    
    header_html = f"""
    <div class="ca-header">
        {logo_html}
        <div class="ca-header-text">
            <h1>ClaimAngel 1160 Reconciliation</h1>
            <p>Upload your bank + repayments file, we'll do the rest.</p>
        </div>
    </div>
    """
    
    st.markdown(header_html, unsafe_allow_html=True)

def get_logo_base64():
    """Convert logo to base64 for embedding"""
    import base64
    logo_path = Path(__file__).parent / "assets" / "claimangel_logo.png"
    try:
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

def create_status_pill(text, pill_type="info"):
    """Create a styled status pill"""
    return f'<span class="pill pill-{pill_type}">{text}</span>'

def show_pre_reconciliation_balance():
    """Show pre-reconciliation balance verification to determine transfer amount"""
    st.markdown('<div class="ca-card" style="margin-top: 2rem;">', unsafe_allow_html=True)
    st.subheader("💰 Pre-Reconciliation Balance Check")
    st.markdown("**Determine how much to transfer out of 1160 account**")
    st.markdown("Compare your Chase balance against expected exclusions to predict the bank revenue.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("**💳 Current Chase Balance**")
        pre_balance = st.number_input(
            "Enter your Chase 1160 balance:",
            value=0.0,
            format="%.2f",
            step=0.01,
            key="pre_balance",
            help="Current balance in your Chase 1160 account"
        )
    
    with col2:
        st.markdown("**➖ Expected Exclusions**")
        st.markdown("*Amounts that don't apply to this reconciliation period*")
        
        # Initialize session state for pre-reconciliation exclusions
        if 'pre_exclusions' not in st.session_state:
            st.session_state.pre_exclusions = []
        
        # Add new exclusion
        new_exclusion = st.number_input(
            "Add exclusion amount:",
            value=0.0,
            format="%.2f",
            step=0.01,
            key="new_pre_exclusion",
            help="Amount to exclude from this reconciliation period"
        )
        
        col2a, col2b = st.columns([1, 1])
        with col2a:
            if st.button("➕ Add", use_container_width=True, key="add_pre_exclusion"):
                if new_exclusion != 0:
                    st.session_state.pre_exclusions.append(new_exclusion)
                    st.rerun()
        
        with col2b:
            if st.button("🗑️ Clear All", use_container_width=True, key="clear_pre_exclusions"):
                st.session_state.pre_exclusions = []
                st.rerun()
    
    # Show current exclusions
    if st.session_state.pre_exclusions:
        st.markdown("**Current Exclusions:**")
        total_exclusions = 0
        for i, exclusion in enumerate(st.session_state.pre_exclusions):
            col_excl1, col_excl2 = st.columns([3, 1])
            with col_excl1:
                st.write(f"${exclusion:,.2f}")
            with col_excl2:
                if st.button("❌", key=f"remove_pre_{i}"):
                    st.session_state.pre_exclusions.pop(i)
                    st.rerun()
            total_exclusions += exclusion
        
        st.markdown(f"**Total Exclusions: ${total_exclusions:,.2f}**")
    else:
        total_exclusions = 0
        st.info("No exclusions added yet")
    
    # Calculate expected transfer amount
    if pre_balance > 0:
        expected_revenue = pre_balance - total_exclusions
        
        st.markdown("---")
        st.markdown("**📊 Expected Transfer Calculation:**")
        
        # Create calculation display
        calc_html = f"""
        <div style="padding: 1.5rem; background: #f8f9fa; border-radius: 12px; margin: 1rem 0;">
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; text-align: center;">
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Chase 1160 Balance</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">${pre_balance:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Minus Exclusions</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">-${total_exclusions:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Expected Revenue</div>
                    <div style="font-size: 1.2rem; font-weight: 600; color: var(--accent);">${expected_revenue:,.2f}</div>
                </div>
            </div>
        </div>
        """
        st.markdown(calc_html, unsafe_allow_html=True)
        
        # Guidance
        if expected_revenue > 0:
            st.success(f"💡 **Expected Bank Revenue: ${expected_revenue:,.2f}**")
            st.info(f"🎯 **Action:** After reconciliation, you should be able to transfer **${expected_revenue:,.2f}** out of the 1160 account if the reconciliation matches perfectly.")
        else:
            st.warning("⚠️ **Warning:** Expected revenue is negative or zero. Please check your exclusions.")
            
        # Store in session state for comparison after reconciliation
        st.session_state.expected_transfer_amount = expected_revenue
    
    st.markdown('</div>', unsafe_allow_html=True)

def show_balance_reconciliation(output_path):
    """Show balance reconciliation interface"""
    st.markdown('<div class="ca-card" style="margin-top: 2rem;">', unsafe_allow_html=True)
    st.subheader("💰 Post-Reconciliation Balance Verification")
    st.markdown("Compare your actual results against the pre-reconciliation estimate.")
    
    # Try to read the Bank Revenue Summary from the Excel file
    bank_revenue_total = 0
    try:
        bank_summary_df = pd.read_excel(output_path, sheet_name='Bank_Revenue_Summary')
        if not bank_summary_df.empty and 'Total_Bank_Based_Revenue' in bank_summary_df.columns:
            bank_revenue_total = float(bank_summary_df['Total_Bank_Based_Revenue'].iloc[0])
    except Exception as e:
        st.warning(f"Could not read Bank Revenue Summary: {e}")
    
    # Check if we have a pre-reconciliation estimate
    expected_amount = getattr(st.session_state, 'expected_transfer_amount', None)
    
    if expected_amount is not None:
        # Compare pre-reconciliation estimate with actual result
        estimate_delta = bank_revenue_total - expected_amount
        
        st.markdown("**📊 Pre vs Post Reconciliation Comparison:**")
        
        comparison_html = f"""
        <div style="padding: 1.5rem; background: #f8f9fa; border-radius: 12px; margin: 1rem 0;">
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; text-align: center;">
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Pre-Reconciliation Estimate</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">${expected_amount:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Actual Bank Revenue</div>
                    <div style="font-size: 1.2rem; font-weight: 600; color: var(--accent);">${bank_revenue_total:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Delta</div>
                    <div style="font-size: 1.2rem; font-weight: 600; color: {'green' if abs(estimate_delta) < 0.01 else 'orange' if abs(estimate_delta) < 100 else 'red'};">
                        ${estimate_delta:+,.2f}
                    </div>
                </div>
            </div>
        </div>
        """
        st.markdown(comparison_html, unsafe_allow_html=True)
        
        if abs(estimate_delta) < 0.01:
            st.success("🎯 **Perfect Prediction!** Your pre-reconciliation estimate was spot on!")
            st.success(f"💰 **Transfer Amount: ${bank_revenue_total:,.2f}** - You can confidently transfer this amount out of the 1160 account.")
        elif abs(estimate_delta) < 100:
            st.warning(f"⚠️ **Small variance of ${abs(estimate_delta):,.2f}** - Close estimate but review for minor discrepancies.")
            st.info(f"💰 **Transfer Amount: ${bank_revenue_total:,.2f}** - This is the actual amount you can transfer.")
        else:
            st.error(f"❌ **Significant variance of ${abs(estimate_delta):,.2f}** - Review your exclusions or there may be unexpected transactions.")
            st.info(f"💰 **Transfer Amount: ${bank_revenue_total:,.2f}** - Use this actual amount, but investigate the variance.")
    else:
        # Display the total bank revenue
        st.markdown(f"""
        <div style="padding: 1rem; background: var(--sky); border-radius: 12px; border: 1px solid var(--accent-light); margin-bottom: 1rem;">
            <strong>📊 Total Bank-based Revenue: ${bank_revenue_total:,.2f}</strong>
        </div>
        """, unsafe_allow_html=True)
        st.info("💡 Use the pre-reconciliation balance check above before running future reconciliations for better predictions!")
    
    # Final transfer recommendation
    st.markdown("---")
    st.markdown(f"""
    <div style="padding: 1.5rem; background: var(--sky); border-radius: 12px; border: 1px solid var(--accent-light); text-align: center;">
        <h3 style="color: var(--teal-dark); margin: 0;">🎯 Final Transfer Recommendation</h3>
        <h2 style="color: var(--accent); margin: 0.5rem 0;">${bank_revenue_total:,.2f}</h2>
        <p style="margin: 0; color: var(--ink);">This is the amount you can safely transfer out of your 1160 account.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("**💳 Optional: Manual Balance Verification**")
        chase_balance = st.number_input(
            "Enter your Chase balance:",
            value=0.0,
            format="%.2f",
            step=0.01,
            key="chase_balance"
        )
    
    with col2:
        st.markdown("**➖ Deductions**")
        
        # Initialize session state for deductions
        if 'deductions' not in st.session_state:
            st.session_state.deductions = []
        
        # Add new deduction
        new_deduction = st.number_input(
            "Add deduction:",
            value=0.0,
            format="%.2f",
            step=0.01,
            key="new_deduction"
        )
        
        col2a, col2b = st.columns([1, 1])
        with col2a:
            if st.button("➕ Add", use_container_width=True):
                if new_deduction != 0:
                    st.session_state.deductions.append(new_deduction)
                    st.rerun()
        
        with col2b:
            if st.button("🗑️ Clear All", use_container_width=True):
                st.session_state.deductions = []
                st.rerun()
    
    # Show current deductions
    if st.session_state.deductions:
        st.markdown("**Current Deductions:**")
        total_deductions = 0
        for i, deduction in enumerate(st.session_state.deductions):
            col_ded1, col_ded2 = st.columns([3, 1])
            with col_ded1:
                st.write(f"${deduction:,.2f}")
            with col_ded2:
                if st.button("❌", key=f"remove_{i}"):
                    st.session_state.deductions.pop(i)
                    st.rerun()
            total_deductions += deduction
        
        st.markdown(f"**Total Deductions: ${total_deductions:,.2f}**")
    else:
        total_deductions = 0
        st.info("No deductions added yet")
    
    # Calculate the reconciliation
    if chase_balance > 0:
        calculated_balance = chase_balance - total_deductions
        delta = calculated_balance - bank_revenue_total
        
        st.markdown("---")
        st.markdown("**📊 Reconciliation Results:**")
        
        # Create results display
        results_html = f"""
        <div style="padding: 1.5rem; background: #f8f9fa; border-radius: 12px; margin: 1rem 0;">
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; text-align: center;">
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Chase Balance</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">${chase_balance:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Minus Deductions</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">-${total_deductions:,.2f}</div>
                </div>
                <div>
                    <div style="color: #6c757d; font-size: 0.9rem; margin-bottom: 0.5rem;">Net Balance</div>
                    <div style="font-size: 1.2rem; font-weight: 600;">${calculated_balance:,.2f}</div>
                </div>
            </div>
        </div>
        """
        st.markdown(results_html, unsafe_allow_html=True)
        
        # Show comparison
        if abs(delta) < 0.01:  # Within 1 cent
            st.success(f"✅ **Perfect Match!** Your balance matches the bank revenue exactly.")
        else:
            delta_color = "#d4bd3b" if abs(delta) < 100 else "#dc3545"
            delta_icon = "⚠️" if abs(delta) < 100 else "❌"
            direction = "higher" if delta > 0 else "lower"
            
            st.markdown(f"""
            <div style="padding: 1rem; background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 12px; margin: 1rem 0;">
                <div style="color: {delta_color}; font-weight: 600;">
                    {delta_icon} <strong>Delta: ${abs(delta):,.2f}</strong>
                </div>
                <div style="color: #856404; margin-top: 0.5rem;">
                    Your net balance is ${abs(delta):,.2f} {direction} than the bank-based revenue.
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Suggestions
            if delta > 0:
                st.info("💡 **Suggestion:** You may have additional deductions to account for, or there could be unrecorded transactions.")
            else:
                st.info("💡 **Suggestion:** You may have missed some credits, or there could be timing differences in the reconciliation.")
    
    st.markdown('</div>', unsafe_allow_html=True)

def main():
    """Main application interface"""
    # Inject custom CSS
    inject_custom_css()
    
    # Show header
    show_header()
    
    # Main upload and configuration card
    st.markdown('<div class="ca-card">', unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 2])
    
    with col1:
        st.subheader("📁 Upload Your File")
        uploaded_file = st.file_uploader(
            "Drag and drop your Excel file here",
            type=['xlsx'],
            help="Upload an Excel file containing both your repayments and bank data sheets"
        )
        
        if uploaded_file:
            file_size = len(uploaded_file.getvalue()) / 1024 / 1024  # MB
            st.markdown(f"""
            <div style="margin-top: 1rem; padding: 1rem; background: var(--sky); border-radius: 12px; border: 1px solid var(--accent-light);">
                <strong>✅ File Ready:</strong> {uploaded_file.name}<br>
                <small>Size: {file_size:.1f} MB</small>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="margin-top: 1rem; padding: 1.5rem; background: #f8f9fa; border-radius: 12px; border: 2px dashed var(--sand); text-align: center;">
                <p style="margin: 0; color: #6c757d;">
                    📄 <strong>Expected format:</strong> Excel file (.xlsx)<br>
                    <small>Should contain separate sheets for repayments and bank data</small>
                </p>
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        st.subheader("⚙️ Configuration")
        
        r2d_sheet = st.text_input(
            "Repayments Sheet Name",
            value="Repayments to Date",
            help="Name of the sheet containing repayment data"
        )
        
        chase_sheet = st.text_input(
            "Bank Sheet Name", 
            value="Chase",
            help="Name of the sheet containing bank transaction data"
        )
        
        st.markdown("---")
        
        use_date_filter = st.checkbox("🗓️ Filter by cutoff date")
        ignore_debits_before = None
        
        if use_date_filter:
            cutoff_date = st.date_input(
                "Ignore debits before:",
                help="Only transactions on or after this date will be processed"
            )
            if cutoff_date:
                ignore_debits_before = cutoff_date.strftime("%Y-%m-%d")
                st.markdown(
                    create_status_pill(f"Cutoff: {cutoff_date.strftime('%m/%d/%Y')}", "info"),
                    unsafe_allow_html=True
                )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Pre-reconciliation Balance Verification
    show_pre_reconciliation_balance()
    
    # Action button
    st.markdown('<div style="text-align: center; margin: 2rem 0;">', unsafe_allow_html=True)
    
    if st.button("🚀 Run Reconciliation", use_container_width=True, type="primary"):
        if uploaded_file is None:
            st.warning("⚠️ Please upload an Excel file before running the reconciliation.")
        else:
            run_reconciliation(uploaded_file, r2d_sheet, chase_sheet, ignore_debits_before)
    
    st.markdown('</div>', unsafe_allow_html=True)

def run_reconciliation(uploaded_file, r2d_sheet, chase_sheet, ignore_debits_before):
    """Run the reconciliation process with modern progress tracking"""
    
    # Create status container
    status_container = st.container()
    progress_container = st.container()
    
    with status_container:
        st.markdown('<div class="ca-card">', unsafe_allow_html=True)
        st.subheader("🔄 Processing Reconciliation")
        
        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        step_pills = st.empty()
    
    try:
        # Step 1: Save uploaded file
        with status_text:
            st.markdown("**Step 1:** Processing uploaded file...")
        with step_pills:
            st.markdown(create_status_pill("Uploading", "info"), unsafe_allow_html=True)
        progress_bar.progress(15)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
            tmp_input.write(uploaded_file.getvalue())
            input_path = tmp_input.name
        
        # Step 2: Prepare output
        with status_text:
            st.markdown("**Step 2:** Preparing output configuration...")
        with step_pills:
            st.markdown(
                create_status_pill("Uploaded", "success") + 
                create_status_pill("Configuring", "info"),
                unsafe_allow_html=True
            )
        progress_bar.progress(30)
        
        timestamp = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"Repayments_to_Date_recon-{timestamp}.xlsx"
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
            output_path = tmp_output.name
        
        # Step 3: Run reconciliation
        with status_text:
            st.markdown("**Step 3:** Running reconciliation algorithm...")
        with step_pills:
            st.markdown(
                create_status_pill("Uploaded", "success") + 
                create_status_pill("Configured", "success") + 
                create_status_pill("Processing", "info"),
                unsafe_allow_html=True
            )
        progress_bar.progress(50)
        
        # Wrap the reconciliation call in a spinner
        with st.spinner("Analyzing data and performing reconciliation..."):
            r2d_recon.run(
                file_path=input_path,
                r2d_sheet=r2d_sheet,
                chase_sheet=chase_sheet, 
                out_path=output_path,
                ignore_debits_before=ignore_debits_before
            )
        
        progress_bar.progress(85)
        
        # Step 4: Finalize
        with status_text:
            st.markdown("**Step 4:** Finalizing results...")
        with step_pills:
            st.markdown(
                create_status_pill("Uploaded", "success") + 
                create_status_pill("Configured", "success") + 
                create_status_pill("Processed", "success") + 
                create_status_pill("Finalizing", "info"),
                unsafe_allow_html=True
            )
        progress_bar.progress(100)
        
        # Read output for download
        with open(output_path, 'rb') as f:
            output_data = f.read()
        
        # Success state
        with status_text:
            st.markdown("**✅ Reconciliation Complete!**")
        with step_pills:
            st.markdown(
                create_status_pill("Uploaded", "success") + 
                create_status_pill("Configured", "success") + 
                create_status_pill("Processed", "success") + 
                create_status_pill("Ready", "success"),
                unsafe_allow_html=True
            )
        
        # Success message and download
        st.success("🎉 Your reconciliation has been completed successfully!")
        
        # File info
        file_size = len(output_data) / 1024 / 1024  # MB
        st.markdown(f"""
        <div style="margin: 1rem 0; padding: 1rem; background: var(--sky); border-radius: 12px; border: 1px solid var(--accent-light);">
            <strong>📊 Results Summary:</strong><br>
            • Output file: {output_filename}<br>
            • File size: {file_size:.1f} MB<br>
            • Generated: {datetime.now().strftime('%I:%M %p on %B %d, %Y')}
        </div>
        """, unsafe_allow_html=True)
        
        # Download button
        st.download_button(
            label="� Download Reconciliation Results",
            data=output_data,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Show balloons animation
        st.balloons()
        
        st.markdown('</div>', unsafe_allow_html=True)
        
    except Exception as e:
        # Error handling
        with status_text:
            st.markdown("**❌ Error Occurred**")
        with step_pills:
            st.markdown(create_status_pill("Error", "warning"), unsafe_allow_html=True)
        
        st.error(f"An error occurred during reconciliation: {str(e)}")
        
        # Detailed error for debugging
        with st.expander("🔍 Technical Details (for debugging)"):
            st.code(traceback.format_exc())
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    finally:
        # Clean up temporary files
        try:
            if 'input_path' in locals():
                os.unlink(input_path)
            if 'output_path' in locals():
                os.unlink(output_path)
        except Exception:
            pass

if __name__ == "__main__":
    # Configure page
    st.set_page_config(
        page_title="ClaimAngel Reconciliation",
        page_icon="assets/favicon.png",
        layout="centered",
        initial_sidebar_state="collapsed"
    )
    
    # Run main app
    main()
