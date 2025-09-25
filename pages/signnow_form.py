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
    st.subheader("Contract Information")
    
    # Document type selection
    st.subheader("📄 Document Selection")
    document_types = {
        "Development Contract": "development_contract",
        "Development Terms and Conditions": "development_terms", 
        "Terms and Conditions": "terms_conditions",
        "Production Contract": "production_contract"
    }
    
    selected_document = st.selectbox(
        "Select Document Type",
        options=list(document_types.keys()),
        help="Choose the type of document to prepare for electronic signature"
    )
    
    template_type = document_types[selected_document]
    
    # Show document description
    document_descriptions = {
        "Development Contract": "Professional development services agreement with contract amount and signature fields",
        "Development Terms and Conditions": "Terms and conditions for development services with signature fields",
        "Terms and Conditions": "General terms and conditions agreement with signature fields", 
        "Production Contract": "Production services agreement (Note: Detailed configuration pending)"
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
    st.subheader("📝 Client Information")
    
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
        # Contract amount field (only show for contract types)
        contract_amount = None
        if template_type in ['development_contract', 'production_contract']:
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
    
    # Convert date to string format if provided
    contract_date_str = None
    if contract_date:
        contract_date_str = contract_date.strftime("%B %d, %Y")
    
    # Validation
    required_fields = [client_name, email]
    if template_type in ['development_contract', 'production_contract']:
        required_fields.append(contract_amount)
    
    if not all(required_fields):
        missing_fields = []
        if not client_name:
            missing_fields.append("Client Name")
        if not email:
            missing_fields.append("Email Address")
        if template_type in ['development_contract', 'production_contract'] and not contract_amount:
            missing_fields.append("Contract Amount")
        
        st.warning(f"⚠️ Please fill in all required fields: {', '.join(missing_fields)}")
        return
    
    # Email validation
    if "@" not in email or "." not in email:
        st.error("❌ Please enter a valid email address.")
        return
    
    # Contract amount validation (only for contract types)
    if template_type in ['development_contract', 'production_contract']:
        try:
            # Remove $ and commas for validation
            amount_str = contract_amount.replace('$', '').replace(',', '')
            float(amount_str)
        except ValueError:
            st.error("❌ Please enter a valid contract amount (e.g., $10,000).")
            return
    
    # Document Preview Section
    st.subheader("📄 Document Preview")
    st.warning("IMAGE/LOGO NOT DISPLAYED IN PREVIEW")
    
    # Create a preview of the document with populated fields
    if st.button("👁️ Preview Document", type="secondary", use_container_width=False):
        with st.spinner("Generating document preview..."):
            try:
                # Process the document to show preview
                from docx_template_processor import DocxTemplateProcessor
                processor = DocxTemplateProcessor()
                
                preview_path = processor.process_document(
                    template_type=template_type,
                    client_name=client_name,
                    email=email,
                    contract_amount=contract_amount,
                    contract_date=contract_date_str
                )
                
                # Build highlight values: only our inputs should be bold
                highlight_values = [v for v in [client_name, email] if v]
                if contract_amount:
                    # Include raw and formatted amount for robust matching
                    amt_clean = contract_amount.replace('$', '').replace(',', '')
                    try:
                        amt_formatted = f"${float(amt_clean):,.2f}"
                        highlight_values.append(amt_formatted)
                    except Exception:
                        pass
                    highlight_values.append(contract_amount)
                if contract_date_str:
                    highlight_values.append(contract_date_str)

                # Convert .docx to PDF for preview, bold only our inputs
                pdf_content = signnow_api._convert_docx_to_pdf(preview_path, highlight_values=highlight_values)
                
                
                # Create a PDF viewer using base64 encoding
                import base64
                pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')
                
                # Display PDF using HTML embed
                pdf_display = f"""
                <iframe src="data:application/pdf;base64,{pdf_base64}" 
                        width="100%" 
                        height="600" 
                        type="application/pdf">
                </iframe>
                """
                st.markdown(pdf_display, unsafe_allow_html=True)
                
                # Store the preview path for sending
                st.session_state['preview_document_path'] = preview_path
                st.session_state['document_ready'] = True
                
            except Exception as e:
                st.error(f"❌ Error generating preview: {str(e)}")
                st.session_state['document_ready'] = False
    
    # Show preview status
    if 'document_ready' in st.session_state and st.session_state['document_ready']:
        st.info("📄 Document is ready for sending!")
    
    # Action button
    st.subheader("🚀 Actions")
    
    # Send Contract button (only enabled if document is ready)
    send_disabled = not st.session_state.get('document_ready', False)
    
    if st.button("📤 Send Contract", type="primary", use_container_width=False, disabled=send_disabled):
        if not send_disabled:
            with st.spinner("Sending contract for electronic signature..."):
                try:
                    # Use the preview document path
                    preview_path = st.session_state.get('preview_document_path')
                    
                    if preview_path and os.path.exists(preview_path):
                        # Build highlight values again for sending
                        highlight_values = [v for v in [client_name, email] if v]
                        if contract_amount:
                            amt_clean = contract_amount.replace('$', '').replace(',', '')
                            try:
                                amt_formatted = f"${float(amt_clean):,.2f}"
                                highlight_values.append(amt_formatted)
                            except Exception:
                                pass
                            highlight_values.append(contract_amount)
                        if contract_date_str:
                            highlight_values.append(contract_date_str)

                        # Send the ORIGINAL .docx file (not PDF) to preserve 100% exact formatting
                        with open(preview_path, 'rb') as docx_file:
                            docx_content = docx_file.read()

                        # Ensure we're authenticated
                        if not getattr(signnow_api, 'access_token', None):
                            if not signnow_api.authenticate():
                                st.error("❌ Failed to authenticate with SignNow. Check credentials.")
                                return
                        
                        # Upload original .docx to SignNow (preserves images, logos, tables, exact formatting)
                        # Use the actual filename from the processed document
                        actual_filename = os.path.basename(preview_path)
                        files = {
                            'file': (actual_filename, docx_content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                        }
                        
                        data = {
                            'name': f'{selected_document} - {client_name}'
                        }
                        
                        headers = {
                            "Authorization": f"Bearer {signnow_api.access_token}"
                        }
                        
                        # Create document in SignNow
                        response = requests.post(
                            f"{signnow_api.base_url}/document",
                            files=files,
                            data=data,
                            headers=headers,
                            timeout=90
                        )
                        if response.status_code >= 400:
                            try:
                                err_text = response.text
                            except Exception:
                                err_text = "<no body>"
                            st.error(f"❌ Failed to create document in SignNow (HTTP {response.status_code}): {err_text}")
                            return
                        
                        create_response = response.json()
                        document_id = create_response.get("id")
                        
                        if document_id:
                            # Send for signing
                            send_url = f"{signnow_api.base_url}/document/{document_id}/invite"
                            send_data = {
                                "to": email,
                                "from": signnow_api.user_email
                            }
                            
                            send_response = requests.post(send_url, json=send_data, headers=headers)
                            send_response.raise_for_status()
                            
                            st.success(f"✅ Original .docx contract sent successfully to {email}!")
                            st.success(f"📄 Document ID: {document_id}")
                            st.balloons()
                            
                            # Clean up processed document to avoid file accumulation
                            try:
                                os.remove(preview_path)
                                print("🗑️ Processed document cleaned up successfully")
                            except Exception as e:
                                st.warning(f"⚠️ Could not clean up processed document: {str(e)}")
                            
                            st.session_state['document_ready'] = False
                            if 'preview_document_path' in st.session_state:
                                del st.session_state['preview_document_path']
                        else:
                            st.error("❌ Failed to create document in SignNow")
                    else:
                        st.error("❌ Preview document not found. Please generate preview first.")
                        
                except Exception as e:
                    st.error(f"❌ Error sending contract: {str(e)}")
                    
                    # Clean up processed document even on error to avoid accumulation
                    try:
                        if 'preview_document_path' in st.session_state:
                            preview_path = st.session_state['preview_document_path']
                            if os.path.exists(preview_path):
                                os.remove(preview_path)
                                print("🗑️ Processed document cleaned up after error")
                    except Exception as cleanup_error:
                        st.warning(f"⚠️ Could not clean up processed document: {str(cleanup_error)}")
        else:
            st.warning("⚠️ Please generate document preview first")
    
    if send_disabled:
        st.caption("💡 Generate document preview first to enable sending")

if __name__ == "__main__":
    main()
