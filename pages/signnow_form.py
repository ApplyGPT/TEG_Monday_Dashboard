"""
SignNow Contract Form Page
Allows users to review and send contracts for signing
"""

import streamlit as st
import sys
import os
import requests

# Add parent directory to path to import signnow_integration
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from signnow_integration import SignNowAPI, load_signnow_credentials
import html

def get_decoded_query_params():
    """Get query parameters and decode HTML entities"""
    query_params = st.query_params
    decoded_params = {}
    
    for key, value in query_params.items():
        # Decode HTML entities (like &amp; -> &)
        decoded_value = html.unescape(str(value))
        decoded_params[key] = decoded_value
    
    return decoded_params

# Page configuration
st.set_page_config(
    page_title="SignNow Contract Form",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to hide all non-tool pages from sidebar navigation
st.markdown("""
<style>
/* Hide ALL sidebar list items by default */
[data-testid="stSidebarNav"] li {
    display: none !important;
}

/* Show list items that contain allowed tool pages using :has() selector */
[data-testid="stSidebarNav"] li:has(a[href*="quickbooks"]),
[data-testid="stSidebarNav"] li:has(a[href*="signnow"]),
[data-testid="stSidebarNav"] li:has(a[href*="/tools"]),
[data-testid="stSidebarNav"] li:has(a[href*="workbook"]) {
    display: block !important;
}
</style>
<script>
// JavaScript to show only tool pages and hide everything else
(function() {
    function showToolPagesOnly() {
        const navItems = document.querySelectorAll('[data-testid="stSidebarNav"] li');
        const allowedPages = ['quickbooks', 'signnow', 'tools', 'workbook'];
        
        navItems.forEach(item => {
            const link = item.querySelector('a');
            if (!link) {
                item.style.setProperty('display', 'none', 'important');
                return;
            }
            
            const href = (link.getAttribute('href') || '').toLowerCase();
            const text = link.textContent.trim().toLowerCase();
            
            // Check if this is an allowed tool page
            const isToolPage = allowedPages.some(page => {
                return href.includes(page) || text.includes(page.toLowerCase());
            });
            
            // Make sure it's not ads dashboard or other dashboards
            const isDashboard = (text.includes('ads') && text.includes('dashboard')) || 
                              (href.includes('ads') && href.includes('dashboard'));
            
            if (isToolPage && !isDashboard) {
                item.style.setProperty('display', 'block', 'important');
                link.style.setProperty('display', 'block', 'important');
            } else {
                item.style.setProperty('display', 'none', 'important');
            }
        });
    }
    
    // Run immediately and on load
    showToolPagesOnly();
    window.addEventListener('load', function() {
        setTimeout(showToolPagesOnly, 50);
        setTimeout(showToolPagesOnly, 200);
        setTimeout(showToolPagesOnly, 500);
    });
    
    // Watch for DOM changes
    const observer = new MutationObserver(function() {
        showToolPagesOnly();
    });
    
    setTimeout(function() {
        const sidebar = document.querySelector('[data-testid="stSidebarNav"]');
        if (sidebar) {
            observer.observe(sidebar, { 
                childList: true, 
                subtree: true,
                attributes: true
            });
        }
    }, 100);
})();
</script>
""", unsafe_allow_html=True)

def main():
    """Main SignNow form function"""
    st.title("📝 SignNow Contract Form")
    st.markdown("Review and send contracts for signing")
    
    # Account selection options
    account_options = {
        "Heather": "heather",
        "Jennifer": "jennifer",
        "Anthony": "anthony"
    }
    
    # Initialize session state for selected account if not exists
    if 'signnow_selected_account' not in st.session_state:
        st.session_state.signnow_selected_account = "Heather"
    
    # Account selection dropdown - moved to main page
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_account_display = st.selectbox(
            "Select Account",
            options=list(account_options.keys()),
            index=list(account_options.keys()).index(st.session_state.signnow_selected_account) if st.session_state.signnow_selected_account in account_options else 0,
            help="Choose which SignNow account to use for sending documents",
            key="signnow_account_selector"
        )
    
    with col2:
        st.markdown("<div style='height: 38px'></div>", unsafe_allow_html=True)  # Align with selectbox
    
    # Update session state when account changes
    st.session_state.signnow_selected_account = selected_account_display
    selected_account = account_options[selected_account_display]
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ SignNow Settings")
        
        # Test connection button
        if st.button("🔗 Test SignNow Connection"):
            with st.spinner("Testing connection..."):
                # Reload credentials with selected account
                account_credentials = load_signnow_credentials(account_name=selected_account)
                if account_credentials:
                    test_api = SignNowAPI(
                        client_id=account_credentials['client_id'],
                        client_secret=account_credentials['client_secret'],
                        basic_auth_token=account_credentials['basic_auth_token'],
                        username=account_credentials['username'],
                        password=account_credentials['password'],
                        api_key=account_credentials.get('api_key')
                    )
                    if test_api.authenticate():
                        st.success("✅ Connection successful!")
                    else:
                        st.error("❌ Connection failed!")
                else:
                    st.error("❌ Could not load credentials for selected account")
    
    # Load SignNow credentials with selected account
    credentials = load_signnow_credentials(account_name=selected_account)
    if not credentials:
        st.error("SignNow credentials not configured. Please check your secrets.toml file.")
        st.stop()
    
    # Initialize SignNow API with selected account credentials
    signnow_api = SignNowAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        basic_auth_token=credentials['basic_auth_token'],
        username=credentials['username'],
        password=credentials['password'],
        api_key=credentials.get('api_key')
    )
    
    # Main form
    
    # Document type selection
    st.subheader("Document Selection")
    document_types = {
        "Development Contract + Development Terms": "development_pair",
        "Production Contract + Production Terms": "production_pair"
    }
    
    selected_document = st.selectbox(
        "Select Document Package",
        options=list(document_types.keys()),
        help="Choose the document package to prepare for electronic signature"
    )
    
    template_type = document_types[selected_document]
    
    # Get data from session state (from redirect) or URL parameters
    if 'signnow_data' in st.session_state:
        # Data from redirect page
        data = st.session_state['signnow_data']
        client_name_default = f"{data.get('first_name', '')} {data.get('last_name', '')}".strip()
        email_default = data.get('email', '')
        contract_amount_default = data.get('contract_amount', '')
        company_name_default = (
            data.get('company_name')
            or data.get('client_company_name')
            or data.get('client_company')
            or ''
        )
        st.success("✅ Data loaded from Monday.com")
    else:
        # Fallback to URL parameters
        query_params = get_decoded_query_params()
        first_name_default = query_params.get('first_name', '')
        last_name_default = query_params.get('last_name', '')
        client_name_default = f"{first_name_default} {last_name_default}".strip()
        email_default = query_params.get('email', '')
        contract_amount_default = query_params.get('contract_amount', '')
        company_name_default = (
            query_params.get('company_name')
            or query_params.get('client_company_name')
            or ''
        )
    
    # Form fields with auto-filled values
    st.subheader("Client Information")
    
    col1, col2 = st.columns(2)
    
    with col1:
        client_name = st.text_input(
            "Client Name",
            value=client_name_default,
            help="Enter the client's full name"
        )
        
        email = st.text_input(
            "Email Address",
            value=email_default,
            help="Enter the client's email address"
        )

        company_name_default = company_name_default.strip()
        company_name_enabled = st.checkbox(
            "Company Name?",
            value=bool(company_name_default),
            help="Check if the client has a company name to include in the documents."
        )

        if company_name_enabled:
            tegmade_for = st.text_input(
                "Client Company Name",
                value=company_name_default,
                key="company_name_input",
                help="Enter the company name to include in the contract."
            )
        else:
            tegmade_for = ""
    
    with col2:
        # Contract amount field (only show for development contracts)
        contract_amount = None
        if template_type == 'development_pair':
            contract_amount = st.text_input(
            "Contract Amount",
            value=contract_amount_default,
            help="Enter the contract amount (e.g., $10,000)"
        )
        
        # Contract date field
        contract_date = st.date_input(
            "Contract Date",
            value=None,
            help="Select the contract date (defaults to current date if not specified)"
        )
    
    # Production Contract specific fields
    if template_type == 'production_pair':
        st.subheader("💰 Production Contract Details")
        
        # Get production parameters from URL or session state
        if 'signnow_data' in st.session_state:
            data = st.session_state['signnow_data']
            total_contract_amount_default = data.get('total_contract_amount', '')
            pre_production_fee_default = data.get('pre_production_fee', '')
            sewing_cost_default = data.get('sewing_cost', '')
            total_due_at_signing_default = data.get('total_due_at_signing', '')
        else:
            # Fallback to URL parameters
            query_params = get_decoded_query_params()
            total_contract_amount_default = query_params.get('total_contract_amount', '')
            pre_production_fee_default = query_params.get('pre_production_fee', '')
            sewing_cost_default = query_params.get('sewing_cost', '')
            total_due_at_signing_default = query_params.get('total_due_at_signing', '')
        
        col3, col4 = st.columns(2)
        
        with col3:
            total_contract_amount = st.text_input(
                "Total Contract Amount", 
                value=total_contract_amount_default,
                help="Enter the total contract amount (e.g., $25,000)"
            )
            
            sewing_cost = st.text_input(
                "Sewing Cost",
                value=sewing_cost_default,
                help="Enter the sewing cost (e.g., $2,500)"
            )
        
        with col4:
            pre_production_fee = st.text_input(
                "Pre-production Fee",
                value=pre_production_fee_default,
                help="Enter the pre-production fee (e.g., $1,000)"
            )
            
            total_due_at_signing = st.text_input(
                "Total Due at Signing",
                value=total_due_at_signing_default,
                help="Enter the total amount due at signing (e.g., $5,000)"
            )
    else:
        # Set production fields to None for non-production contracts
        total_contract_amount = None
        pre_production_fee = None
        sewing_cost = None
        total_due_at_signing = None
    
    uploaded_pdf = st.file_uploader(
        "Upload a PDF of the spreadsheet to be inserted into the contract document",
        type=['pdf'],
        help="Upload a PDF of the spreadsheet that will be inserted after the first paragraph in the contract"
    )
    
    # Show uploaded PDF preview
    if uploaded_pdf is not None:
        st.success("✅ PDF uploaded successfully!")
        # Display PDF preview
        import base64
        pdf_bytes = uploaded_pdf.getvalue()
        pdf_base64 = base64.b64encode(pdf_bytes).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{pdf_base64}" width="100%" height="600" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
    else:
        st.warning("⚠️ Please upload a spreadsheet PDF")
    
    # Convert date to string format if provided
    contract_date_str = None
    if contract_date:
        contract_date_str = contract_date.strftime("%B %d, %Y")
    
    # Validation
    required_fields = [client_name, email, uploaded_pdf]
    if template_type == 'development_pair':
        required_fields.append(contract_amount)
    elif template_type == 'production_pair':
        required_fields.extend([total_contract_amount, pre_production_fee, sewing_cost, total_due_at_signing])
    
    if not all(required_fields):
        missing_fields = []
        if not client_name:
            missing_fields.append("Client Name")
        if not email:
            missing_fields.append("Email Address")
        if not uploaded_pdf:
            missing_fields.append("Spreadsheet PDF")
        if template_type == 'development_pair' and not contract_amount:
            missing_fields.append("Contract Amount")
        if template_type == 'production_pair':
            if not total_contract_amount:
                missing_fields.append("Total Contract Amount")
            if not sewing_cost:
                missing_fields.append("Sewing Cost")
            if not pre_production_fee:
                missing_fields.append("Pre-production Fee")
            if not total_due_at_signing:
                missing_fields.append("Total Due at Signing")
        
        st.warning(f"⚠️ Please fill in all required fields: {', '.join(missing_fields)}")
        return
    
    # Email validation
    if "@" not in email or "." not in email:
        st.error("❌ Please enter a valid email address.")
        return
    
    # Contract amount validation (only for development contracts)
    if template_type == 'development_pair':
        try:
            # Remove $ and commas for validation
            amount_str = contract_amount.replace('$', '').replace(',', '')
            float(amount_str)
        except ValueError:
            st.error("❌ Please enter a valid contract amount (e.g., $10,000).")
            return
    
    # Production contract amount validation
    if template_type == 'production_pair':
        production_amounts = [total_contract_amount, sewing_cost, pre_production_fee, total_due_at_signing]
        for amount_field, field_name in zip(production_amounts, 
                                          ['Total Contract Amount', 'Sewing Cost', 'Pre-production Fee', 'Total Due at Signing']):
            try:
                # Remove $ and commas for validation
                amount_str = amount_field.replace('$', '').replace(',', '')
                float(amount_str)
            except ValueError:
                st.error(f"❌ Please enter a valid {field_name} (e.g., $10,000).")
                return
    
    # Documents Preview Section
    st.subheader("📄 Documents Preview")
    st.warning("IMAGE/LOGO NOT DISPLAYED IN PREVIEW")
    
    # Create previews of both documents with populated fields
    if st.button("👁️ Preview Documents", type="secondary", use_container_width=False):
        with st.spinner("Generating document previews..."):
            try:
                # Process the document(s) to show preview
                from docx_template_processor import DocxTemplateProcessor
                processor = DocxTemplateProcessor()
                
                preview_paths = {}
                
                # Determine which documents to process based on template type
                if template_type == 'development_pair':
                    # Process development contract
                    preview_paths['contract'] = processor.process_document(
                        template_type='development_contract',
                        client_name=client_name,
                        email=email,
                        contract_amount=contract_amount,
                        contract_date=contract_date_str,
                        uploaded_pdf=uploaded_pdf,
                        tegmade_for=tegmade_for
                    )
                    
                    # Process development terms
                    preview_paths['terms'] = processor.process_document(
                        template_type='development_terms',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,
                        contract_date=contract_date_str,
                        tegmade_for=tegmade_for
                    )
                    
                elif template_type == 'production_pair':
                    # Process production contract
                    preview_paths['contract'] = processor.process_document(
                        template_type='production_contract',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,  # Not used for production
                        contract_date=contract_date_str,
                        total_contract_amount=total_contract_amount,
                        sewing_cost=sewing_cost,
                        pre_production_fee=pre_production_fee,
                        total_due_at_signing=total_due_at_signing,
                        uploaded_pdf=uploaded_pdf,
                        tegmade_for=tegmade_for
                    )
                    
                    # Process production terms
                    preview_paths['terms'] = processor.process_document(
                        template_type='production_terms',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,
                        contract_date=contract_date_str,
                        tegmade_for=tegmade_for
                    )
                
                # Build highlight values: only our inputs should be bold
                highlight_values = [v for v in [client_name, email, tegmade_for] if v]
                
                if template_type == 'development_pair' and contract_amount:
                    # Include raw and formatted amount for robust matching
                    amt_clean = contract_amount.replace('$', '').replace(',', '')
                    try:
                        amt_formatted = f"${float(amt_clean):,.2f}"
                        highlight_values.append(amt_formatted)
                    except Exception:
                        pass
                    highlight_values.append(contract_amount)
                
                elif template_type == 'production_pair':
                    # Add all production contract amounts
                    production_amounts = [total_contract_amount, sewing_cost, pre_production_fee, total_due_at_signing]
                    for amount in production_amounts:
                        if amount:
                            amt_clean = amount.replace('$', '').replace(',', '')
                            try:
                                amt_formatted = f"${float(amt_clean):,.2f}"
                                highlight_values.append(amt_formatted)
                            except Exception:
                                pass
                            highlight_values.append(amount)
                
                if contract_date_str:
                    highlight_values.append(contract_date_str)

                # Display both document previews
                import base64
                
                # Show Contract Preview
                st.subheader("📋 Contract Preview")
                contract_pdf_content = signnow_api._convert_docx_to_pdf(preview_paths['contract'], highlight_values=highlight_values)
                contract_pdf_base64 = base64.b64encode(contract_pdf_content).decode('utf-8')
                
                contract_pdf_display = f"""
                <iframe src="data:application/pdf;base64,{contract_pdf_base64}" 
                        width="100%" 
                        height="600" 
                        type="application/pdf">
                </iframe>
                """
                st.markdown(contract_pdf_display, unsafe_allow_html=True)
                
                # Show Terms Preview
                st.subheader("📋 Terms and Conditions Preview")
                terms_pdf_content = signnow_api._convert_docx_to_pdf(preview_paths['terms'], highlight_values=highlight_values)
                terms_pdf_base64 = base64.b64encode(terms_pdf_content).decode('utf-8')
                
                terms_pdf_display = f"""
                <iframe src="data:application/pdf;base64,{terms_pdf_base64}" 
                        width="100%" 
                        height="600" 
                        type="application/pdf">
                </iframe>
                """
                st.markdown(terms_pdf_display, unsafe_allow_html=True)
                
                # Store the preview paths for sending
                st.session_state['preview_document_paths'] = preview_paths
                st.session_state['document_ready'] = True
                
            except Exception as e:
                st.error(f"❌ Error generating preview: {str(e)}")
                st.session_state['document_ready'] = False
    
    # Show preview status
    if 'document_ready' in st.session_state and st.session_state['document_ready']:
        st.info("📄 Documents are ready for sending!")
    
    # Action button
    st.subheader("🚀 Actions")
    
    # Send Documents button (only enabled if documents are ready)
    send_disabled = not st.session_state.get('document_ready', False)
    
    if st.button("📤 Send Documents", type="primary", use_container_width=False, disabled=send_disabled):
        if not send_disabled:
            with st.spinner("Sending documents for electronic signature..."):
                try:
                    # Use the preview document paths
                    preview_paths = st.session_state.get('preview_document_paths')
                    
                    if preview_paths and all(os.path.exists(path) for path in preview_paths.values()):
                        # Ensure we're authenticated
                        if not getattr(signnow_api, 'access_token', None):
                            if not signnow_api.authenticate():
                                st.error("❌ Failed to authenticate with SignNow. Check credentials.")
                                return

                        # Build highlight values similar to preview for better fidelity
                        highlight_values = [v for v in [client_name, email, tegmade_for] if v]
                        if template_type == 'development_pair' and contract_amount:
                            try:
                                amt_clean = contract_amount.replace('$', '').replace(',', '')
                                amt_formatted = f"${float(amt_clean):,.2f}"
                                highlight_values.append(amt_formatted)
                            except Exception:
                                pass
                            highlight_values.append(contract_amount)
                        elif template_type == 'production_pair':
                            production_amounts = [total_contract_amount, sewing_cost, pre_production_fee, total_due_at_signing]
                            for amount in production_amounts:
                                if amount:
                                    amt_clean = amount.replace('$', '').replace(',', '')
                                    try:
                                        amt_formatted = f"${float(amt_clean):,.2f}"
                                        highlight_values.append(amt_formatted)
                                    except Exception:
                                        pass
                                    highlight_values.append(amount)
                        if contract_date_str:
                            highlight_values.append(contract_date_str)

                        # Merge Contract + Terms into a single document and send once
                        # Build filename like: development_contract_terms_Estevao_Cavalcante_October_31_2025
                        if template_type == 'development_pair':
                            prefix = 'development_contract_terms'
                        else:
                            prefix = 'production_contract_terms'
                        client_part = (client_name or '').replace(' ', '_')
                        date_part = (contract_date_str or '').replace(',', '').replace(' ', '_')
                        doc_name = f"{prefix}_{client_part}_{date_part}".strip('_')
                        # Use DOCX merge path to preserve original images/logos/content
                        ok, msg = signnow_api.create_and_send_merged_pair_docx(
                            pair_type=template_type,
                            contract_docx_path=preview_paths['contract'],
                            terms_docx_path=preview_paths['terms'],
                            document_name=doc_name,
                            email=email,
                        )
                        if ok:
                            st.success(f"✅ {msg}")
                            st.balloons()
                        else:
                            st.error(f"❌ {msg}")
                        
                        # Clean up processed documents
                        for path in preview_paths.values():
                            try:
                                os.remove(path)
                            except Exception as e:
                                st.warning(f"⚠️ Could not clean up {os.path.basename(path)}: {str(e)}")
                        
                        # Clean up session state
                        st.session_state['document_ready'] = False
                        if 'preview_document_paths' in st.session_state:
                            del st.session_state['preview_document_paths']
                    else:
                        st.error("❌ Preview documents not found. Please generate preview first.")
                        
                except Exception as e:
                    st.error(f"❌ Error sending documents: {str(e)}")
                    
                    # Clean up processed documents even on error to avoid accumulation
                    try:
                        if 'preview_document_paths' in st.session_state:
                            preview_paths = st.session_state['preview_document_paths']
                            for path in preview_paths.values():
                                if os.path.exists(path):
                                    os.remove(path)
                    except Exception as cleanup_error:
                        st.warning(f"⚠️ Could not clean up processed documents: {str(cleanup_error)}")
    
    if send_disabled:
        st.caption("💡 Generate document previews first to enable sending")

if __name__ == "__main__":
    main()