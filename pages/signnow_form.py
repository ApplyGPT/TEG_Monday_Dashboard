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

def main():
    """Main SignNow form function"""
    st.title("📝 SignNow Contract Form")
    st.markdown("Review and send contracts for signing")
    
    # Load SignNow credentials
    credentials = load_signnow_credentials()
    if not credentials:
        st.error("SignNow credentials not configured. Please check your secrets.toml file.")
        st.stop()
    
    # Initialize SignNow API
    signnow_api = SignNowAPI(
        client_id=credentials['client_id'],
        client_secret=credentials['client_secret'],
        basic_auth_token=credentials['basic_auth_token'],
        username=credentials['username'],
        password=credentials['password'],
        api_key=credentials.get('api_key')
    )
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ SignNow Settings")
        st.info("Configured with SignNow API")
        
        # Test connection button
        if st.button("🔗 Test SignNow Connection"):
            with st.spinner("Testing connection..."):
                if signnow_api.authenticate():
                    st.success("✅ Connection successful!")
                else:
                    st.error("❌ Connection failed!")
    
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
    
    # Show document description
    document_descriptions = {
        "Development Contract + Development Terms": "Development services agreement with terms and conditions - includes contract amount field",
        "Production Contract + Production Terms": "Production services agreement with terms and conditions - includes deposit, total amount, sewing cost, and pre-production fee fields"
    }
    
    st.info(f"📋 **{selected_document}**: {document_descriptions[selected_document]}")
    
    # Get data from session state (from redirect) or URL parameters
    if 'signnow_data' in st.session_state:
        # Data from redirect page
        data = st.session_state['signnow_data']
        client_name_default = f"{data.get('first_name', '')} {data.get('last_name', '')}".strip()
        email_default = data.get('email', '')
        contract_amount_default = data.get('contract_amount', '')
        st.success("✅ Data loaded from Monday.com")
    else:
        # Fallback to URL parameters
        query_params = get_decoded_query_params()
        first_name_default = query_params.get('first_name', '')
        last_name_default = query_params.get('last_name', '')
        client_name_default = f"{first_name_default} {last_name_default}".strip()
        email_default = query_params.get('email', '')
        contract_amount_default = query_params.get('contract_amount', '')
    
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
            deposit_amount_default = data.get('deposit_amount', '')
            total_contract_amount_default = data.get('total_contract_amount', '')
            sewing_cost_default = data.get('sewing_cost', '')
            pre_production_fee_default = data.get('pre_production_fee', '')
        else:
            # Fallback to URL parameters
            query_params = get_decoded_query_params()
            deposit_amount_default = query_params.get('deposit_amount', '')
            total_contract_amount_default = query_params.get('total_contract_amount', '')
            sewing_cost_default = query_params.get('sewing_cost', '')
            pre_production_fee_default = query_params.get('pre_production_fee', '')
        
        col3, col4 = st.columns(2)
        
        with col3:
            deposit_amount = st.text_input(
                "Deposit Amount",
                value=deposit_amount_default,
                help="Enter the deposit amount (e.g., $5,000)"
            )
            
            total_contract_amount = st.text_input(
                "Total Contract Amount", 
                value=total_contract_amount_default,
                help="Enter the total contract amount (e.g., $25,000)"
            )
        
        with col4:
            sewing_cost = st.text_input(
                "Sewing Cost",
                value=sewing_cost_default,
                help="Enter the sewing cost (e.g., $2,500)"
            )
            
            pre_production_fee = st.text_input(
                "Pre-production Fee",
                value=pre_production_fee_default,
                help="Enter the pre-production fee (e.g., $1,000)"
            )
    else:
        # Set production fields to None for non-production contracts
        deposit_amount = None
        total_contract_amount = None
        sewing_cost = None
        pre_production_fee = None
    
    uploaded_image = st.file_uploader(
        "Upload a screenshot of the spreadsheet to be inserted into the contract document",
        type=['png', 'jpg', 'jpeg'],
        help="Upload a screenshot of the spreadsheet that will be inserted after the first paragraph in the contract"
    )
    
    # Show uploaded image preview
    if uploaded_image is not None:
        st.image(uploaded_image, caption="Spreadsheet Screenshot Preview", use_container_width=True)
        st.success("✅ Image uploaded successfully!")
    else:
        st.warning("⚠️ Please upload a spreadsheet screenshot")
    
    # Convert date to string format if provided
    contract_date_str = None
    if contract_date:
        contract_date_str = contract_date.strftime("%B %d, %Y")
    
    # Validation
    required_fields = [client_name, email, uploaded_image]
    if template_type == 'development_pair':
        required_fields.append(contract_amount)
    elif template_type == 'production_pair':
        required_fields.extend([deposit_amount, total_contract_amount, sewing_cost, pre_production_fee])
    
    if not all(required_fields):
        missing_fields = []
        if not client_name:
            missing_fields.append("Client Name")
        if not email:
            missing_fields.append("Email Address")
        if not uploaded_image:
            missing_fields.append("Spreadsheet Screenshot")
        if template_type == 'development_pair' and not contract_amount:
            missing_fields.append("Contract Amount")
        if template_type == 'production_pair':
            if not deposit_amount:
                missing_fields.append("Deposit Amount")
            if not total_contract_amount:
                missing_fields.append("Total Contract Amount")
            if not sewing_cost:
                missing_fields.append("Sewing Cost")
            if not pre_production_fee:
                missing_fields.append("Pre-production Fee")
        
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
        production_amounts = [deposit_amount, total_contract_amount, sewing_cost, pre_production_fee]
        for amount_field, field_name in zip(production_amounts, 
                                          ['Deposit Amount', 'Total Contract Amount', 'Sewing Cost', 'Pre-production Fee']):
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
                        uploaded_image=uploaded_image
                    )
                    
                    # Process development terms
                    preview_paths['terms'] = processor.process_document(
                        template_type='development_terms',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,
                        contract_date=contract_date_str
                    )
                    
                elif template_type == 'production_pair':
                    # Process production contract
                    preview_paths['contract'] = processor.process_document(
                        template_type='production_contract',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,  # Not used for production
                        contract_date=contract_date_str,
                        deposit_amount=deposit_amount,
                        total_contract_amount=total_contract_amount,
                        sewing_cost=sewing_cost,
                        pre_production_fee=pre_production_fee,
                        uploaded_image=uploaded_image
                    )
                    
                    # Process production terms
                    preview_paths['terms'] = processor.process_document(
                        template_type='production_terms',
                        client_name=client_name,
                        email=email,
                        contract_amount=None,
                        contract_date=contract_date_str
                    )
                
                # Build highlight values: only our inputs should be bold
                highlight_values = [v for v in [client_name, email] if v]
                
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
                    production_amounts = [deposit_amount, total_contract_amount, sewing_cost, pre_production_fee]
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
                        
                        document_ids = []
                        
                        # Prepare contract document for upload
                        with open(preview_paths['contract'], 'rb') as contract_file:
                            contract_content = contract_file.read()
                        
                        contract_filename = os.path.basename(preview_paths['contract'])
                        contract_files = {
                            'file': (contract_filename, contract_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        }
                        
                        contract_data = {
                            'name': f'{selected_document} - Contract - {client_name}'
                        }
                        
                        headers = {
                            "Authorization": f"Bearer {signnow_api.access_token}"
                        }
                        
                        # Create contract document in SignNow with retry logic
                        contract_uploaded = False
                        contract_document_id = None
                        
                        for attempt in range(3):  # Try up to 3 times
                            try:
                                if attempt == 0:  # Only show message on first attempt
                                    st.info("📋 Uploading Contract document...")
                                contract_response = requests.post(
                                    f"{signnow_api.base_url}/document",
                                    files=contract_files,
                                    data=contract_data,
                                    headers=headers,
                                    timeout=180  # Increased timeout to 3 minutes
                                )
                                contract_response.raise_for_status()
                                
                                contract_create_response = contract_response.json()
                                contract_document_id = contract_create_response.get("id")
                                contract_uploaded = True
                                st.success("✅ Contract document uploaded successfully!")
                                break
                                
                            except requests.exceptions.Timeout:
                                if attempt < 2:  # Not the last attempt
                                    continue  # Silent retry
                                else:
                                    st.error("❌ Contract upload failed after 3 attempts due to timeout")
                                    raise
                            except Exception as e:
                                st.error(f"❌ Contract upload failed: {str(e)}")
                                raise
                        
                        if not contract_uploaded:
                            st.error("❌ Failed to upload contract document")
                            return
                        
                        # Prepare terms document for upload
                        with open(preview_paths['terms'], 'rb') as terms_file:
                            terms_content = terms_file.read()
                        
                        terms_filename = os.path.basename(preview_paths['terms'])
                        terms_files = {
                            'file': (terms_filename, terms_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        }
                        
                        terms_data = {
                            'name': f'{selected_document} - Terms - {client_name}'
                        }
                        
                        # Create terms document in SignNow with retry logic
                        terms_uploaded = False
                        terms_document_id = None
                        
                        for attempt in range(3):  # Try up to 3 times
                            try:
                                if attempt == 0:  # Only show message on first attempt
                                    st.info("📋 Uploading Terms document...")
                                terms_response = requests.post(
                                    f"{signnow_api.base_url}/document",
                                    files=terms_files,
                                    data=terms_data,
                                    headers=headers,
                                    timeout=180  # Increased timeout to 3 minutes
                                )
                                terms_response.raise_for_status()
                                
                                terms_create_response = terms_response.json()
                                terms_document_id = terms_create_response.get("id")
                                terms_uploaded = True
                                st.success("✅ Terms document uploaded successfully!")
                                break
                                
                            except requests.exceptions.Timeout:
                                if attempt < 2:  # Not the last attempt
                                    continue  # Silent retry
                                else:
                                    st.error("❌ Terms upload failed after 3 attempts due to timeout")
                                    raise
                            except Exception as e:
                                st.error(f"❌ Terms upload failed: {str(e)}")
                                raise
                        
                        if not terms_uploaded:
                            st.error("❌ Failed to upload terms document")
                            return
                        
                        # Send both documents for signing
                        for doc_id, doc_type in [(contract_document_id, "Contract"), (terms_document_id, "Terms")]:
                            send_url = f"{signnow_api.base_url}/document/{doc_id}/invite"
                            send_data = {
                                "to": email,
                                "from": signnow_api.user_email
                            }
                            
                            send_response = requests.post(send_url, json=send_data, headers=headers)
                            send_response.raise_for_status()
                        
                        st.success(f"✅ Both documents sent successfully to {email}!")
                        st.balloons()
                        
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
