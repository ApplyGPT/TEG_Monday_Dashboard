from flask import Flask, request, jsonify, redirect
import json
import openai
import toml
import os
import csv
from datetime import datetime
import urllib.parse

app = Flask(__name__)

# Load OpenAI API key from Streamlit secrets.toml
def load_openai_key():
    """Load OpenAI API key from environment variable or .streamlit/secrets.toml"""
    # 1. Check environment variable
    env_key = os.environ.get("OPENAI_API_KEY")
    if env_key:
        return env_key

    # 2. Check secrets.toml
    try:
        secrets_path = os.path.join('.streamlit', 'secrets.toml')
        if os.path.exists(secrets_path):
            secrets = toml.load(secrets_path)
            return secrets.get('openai', {}).get('api_key')
        else:
            print("⚠️ Warning: secrets.toml not found in .streamlit/ directory.")
            return None
    except Exception as e:
        print(f"❌ Error while loading secrets.toml: {e}")
        return None


openai_api_key = load_openai_key()

if not openai_api_key:
    print("⚠️ OpenAI API key not found in env var or secrets.toml. Using dummy key for testing.")
    openai_api_key = "dummy_key_for_testing"

openai.api_key = openai_api_key

# Qualification prompt template (updated with 0-3 scoring system)
PROMPT_TEMPLATE = """
ANALYZE THE FOLLOWING LEAD INFORMATION FROM JENNIFER'S FASHION MANUFACTURING FORM.
CLASSIFY THE LEAD USING THE NEW 4-LEVEL SCORING SYSTEM:
SCORE 0: SPAM
SCORE 1: NOT RIGHT FIT
SCORE 2: UNSURE
SCORE 3: RIGHT FIT

Lead Information:
- First Name: {first_name}
- Last Name: {last_name}
- Email: {email}
- Phone Number: {phone_number}
- About Project: {about_project}

CRITERIA FOR SCORE 0 (SPAM):
- **Clear spam indicators:** Generic promotional messages, advertising services unrelated to fashion manufacturing, or completely irrelevant inquiries (e.g., "Do you need an electrician?", "SEO services", "Marketing opportunities").
- **Fake or suspicious information:** Obviously fake names, email addresses from disposable mail services, nonsensical phone numbers, or gibberish text.
- **Bot-generated content:** Repetitive, template-like messages with no real substance or context.

CRITERIA FOR SCORE 1 (NOT RIGHT FIT):
- **Primary Reason:** The lead's project is not a right fit for **The Evans Group (TEG)**, a high-end, luxury fashion manufacturer specializing in couture and small-to-medium volume production, even though the inquiry appears legitimate.
- **Specific examples of "Not Right Fit" projects:**
    - Leads requesting large, mass-market quantities (e.g., thousands of units per style) that do not align with TEG's specialized **small-to-medium volume production model (1-300 pieces per style)**.
    - Projects for basic apparel, sports, gymnastics and/or rhythmic uniforms, simple activewear, or workwear.
    - Medical apparel, underwear, or other non-fashion garments.
    - Budget-focused projects that don't align with luxury manufacturing.

CRITERIA FOR SCORE 2 (UNSURE):
- **When to use:** The lead appears legitimate and professional, but there's not enough information to confidently classify as right fit or not right fit.
- **Examples:** 
    - Vague project descriptions that could be legitimate but lack detail about materials, quantities, or design vision.
    - Projects that might align with TEG but need more information to determine fit.
    - Professional inquiries with incomplete information about the scope or nature of the project.

CRITERIA FOR SCORE 3 (RIGHT FIT):
- **Primary Reason:** The "About Project" description is detailed and clearly aligns with TEG's services for **high-end, luxury, and couture fashion brands**. This is a crucial requirement.
- **Project alignment:** The lead's request fits with TEG's services, which include:
    - **Client Types:** Established high-end designers, luxury brands, or emerging designers with a professional vision.
    - **Production Scale:** The request aligns with **small-to-medium volume production (1-300 pieces per style)**, or other TEG services like pattern making or sample creation.
    - **Specific products:** Runway looks, bridal, evening wear, formal wear, high-end menswear and womenswear, complex knitwear, or specific, elevated basics.
    - **Specific materials:** Delicate silks, leather, wovens, beading, lace, or complex knits.
    - **Project Details:** The request includes specific details on materials, quantities, design elements, and timelines.
- **Professionalism:** The lead provides clear, professional information and demonstrates understanding of luxury fashion manufacturing.

OUTPUT FORMAT:
Provide ONLY a JSON response with the following structure, AND NOTHING ELSE:
{{
  "score": 0, 1, 2, or 3,
  "confidence": "high/medium/low",
  "reason": "brief explanation of classification decision"
}}

Focus on the substance, specificity, and professional context of the lead profile, and its alignment with TEG's specialized business model.
"""

# Calendar URLs for redirection based on score
def get_calendar_url(score, lead_data):
    """
    Get the appropriate calendar URL based on score and lead data
    
    Args:
        score (int): Qualification score (0, 1, 2, or 3)
        lead_data (dict): Lead information for pre-filling
        
    Returns:
        str: Calendar URL with pre-filled data
    """
    if score == 0 or score == 1:
        return "https://tegmade.com/thank-you/"
    elif score == 2:
        return generate_calendly_url_lets_chat(lead_data)
    elif score == 3:
        return generate_calendly_url_introductory_call(lead_data)
    else:
        return "https://tegmade.com/thank-you/"  # Default fallback

def log_request(lead_data, result):
    """Log the request and response to a CSV file"""
    try:
        log_file = 'lead_qualification_log.csv'
        file_exists = os.path.exists(log_file)
        
        with open(log_file, 'a', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['timestamp', 'first_name', 'last_name', 'email', 'phone_number', 
                         'about_project', 'score', 'confidence', 'reason']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Write header if file is new
            if not file_exists:
                writer.writeheader()
            
            # Write the log entry
            writer.writerow({
                'timestamp': datetime.now().isoformat(),
                'first_name': lead_data.get('first_name', ''),
                'last_name': lead_data.get('last_name', ''),
                'email': lead_data.get('email', ''),
                'phone_number': lead_data.get('phone_number', ''),
                'about_project': lead_data.get('about_project', ''),
                'score': result.get('score', ''),
                'confidence': result.get('confidence', ''),
                'reason': result.get('reason', '')
            })
    except Exception as e:
        print(f"Error logging request: {e}")

def format_phone_for_calendly_location(phone_number):
    """
    Format phone number for Calendly location parameter in +1XXXXXXXXXX format
    
    Args:
        phone_number (str): Raw phone number input
        
    Returns:
        str: Formatted phone number for location parameter or None if invalid
    """
    if not phone_number:
        return None
    
    # Remove all non-digit characters
    digits_only = ''.join(filter(str.isdigit, phone_number))
    
    # Handle different input formats
    if len(digits_only) == 10:
        # US number without country code: 2064124253 -> +12064124253
        return f"+1{digits_only}"
    elif len(digits_only) == 11 and digits_only.startswith('1'):
        # US number with country code: 12064124253 -> +12064124253
        return f"+{digits_only}"
    elif len(digits_only) > 11:
        # International number: keep as is but add + if not present
        if not phone_number.startswith('+'):
            return f"+{digits_only}"
        return phone_number
    else:
        # Invalid length, return original if it looks like it has formatting
        if any(char in phone_number for char in ['-', '(', ')', ' ', '+']):
            return phone_number
        return None

def format_phone_for_calendly(phone_number):
    """
    Format phone number for Calendly in +1 XXX-XXX-XXXX format
    
    Args:
        phone_number (str): Raw phone number input
        
    Returns:
        str: Formatted phone number or None if invalid
    """
    if not phone_number:
        return None
    
    # Remove all non-digit characters
    digits_only = ''.join(filter(str.isdigit, phone_number))
    
    # Handle different input formats
    if len(digits_only) == 10:
        # US number without country code: 2064124253 -> +1 206-412-4253
        return f"+1 {digits_only[:3]}-{digits_only[3:6]}-{digits_only[6:]}"
    elif len(digits_only) == 11 and digits_only.startswith('1'):
        # US number with country code: 12064124253 -> +1 206-412-4253
        return f"+1 {digits_only[1:4]}-{digits_only[4:7]}-{digits_only[7:]}"
    elif len(digits_only) > 11:
        # International number: keep as is but add + if not present
        if not phone_number.startswith('+'):
            return f"+{digits_only}"
        return phone_number
    else:
        # Invalid length, return original if it looks like it has formatting
        if any(char in phone_number for char in ['-', '(', ')', ' ', '+']):
            return phone_number
        return None

def generate_calendly_url_lets_chat(lead_data):
    """
    Generate Calendly URL for 'let's chat' (score 2) with pre-filled data
    
    Pre-fills: Name (Full Name), Email, Phone Number, Send text messages to (same as Phone Number), 
    and About Project
    """
    base_url = "https://calendly.com/jamie-the-evans-group/teg-let-s-chat"
    
    # Prepare parameters
    params = {}
    
    # Full name
    first_name = lead_data.get('first_name', '')
    last_name = lead_data.get('last_name', '')
    full_name = f"{first_name} {last_name}".strip()
    if full_name:
        params['name'] = full_name
    
    # Email
    email = lead_data.get('email', '')
    if email:
        params['email'] = email
    
    # Phone number (for both phone and text messages)
    phone_number = lead_data.get('phone_number', '')
    if phone_number:
        # Format phone number for Calendly location parameter
        formatted_phone = format_phone_for_calendly_location(phone_number)
        if formatted_phone:
            params['location'] = formatted_phone  # Location parameter for phone number
    
    # About project
    about_project = lead_data.get('about_project', '')
    if about_project:
        params['a2'] = about_project  # About project parameter (second custom question)
    
    # Build URL with parameters
    if params:
        param_string = urllib.parse.urlencode(params)
        return f"{base_url}?{param_string}"
    else:
        return base_url

def generate_calendly_url_introductory_call(lead_data):
    """
    Generate Calendly URL for 'introductory call' (score 3) with pre-filled data
    
    Pre-fills: Name, Email, and About Project
    """
    base_url = "https://calendly.com/d/ctc8-ndq-rjz/teg-introductory-call"
    
    # Prepare parameters
    params = {}
    
    # Full name
    first_name = lead_data.get('first_name', '')
    last_name = lead_data.get('last_name', '')
    full_name = f"{first_name} {last_name}".strip()
    if full_name:
        params['name'] = full_name
    
    # Email
    email = lead_data.get('email', '')
    if email:
        params['email'] = email
    
    # About project
    about_project = lead_data.get('about_project', '')
    if about_project:
        params['a1'] = about_project  # About project parameter (first custom question)
    
    # Build URL with parameters
    if params:
        param_string = urllib.parse.urlencode(params)
        return f"{base_url}?{param_string}"
    else:
        return base_url

def qualify_lead(lead_data):
    """
    Qualify a lead using OpenAI GPT-4o-mini model
    
    Args:
        lead_data (dict): Dictionary containing lead information
        
    Returns:
        dict: Qualification result with score, confidence, and reason
    """
    try:
        # Check if API key is properly configured
        if openai_api_key == "dummy_key_for_testing":
            return {
                "score": 0,
                "confidence": "low",
                "reason": "OpenAI API key not configured. Please set OPENAI_API_KEY environment variable."
            }
        
        # Format the prompt with the simplified lead data
        prompt = PROMPT_TEMPLATE.format(
            first_name=lead_data.get('first_name', ''),
            last_name=lead_data.get('last_name', ''),
            email=lead_data.get('email', ''),
            phone_number=lead_data.get('phone_number', ''),
            about_project=lead_data.get('about_project', '')
        )
        
        # Send the prompt to OpenAI
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        
        # Parse the JSON response
        result = json.loads(response.choices[0].message.content)
        return result
        
    except json.JSONDecodeError as e:
        return {
            "score": 0,
            "confidence": "low",
            "reason": f"Error parsing AI response: {str(e)}"
        }
    except Exception as e:
        return {
            "score": 0,
            "confidence": "low", 
            "reason": f"Error during qualification: {str(e)}"
        }

@app.route('/qualify', methods=['GET'])
def qualify_lead_get_endpoint():
    """
    GET endpoint that accepts query parameters and redirects based on qualification score
    
    Expected query parameters:
    - first_name: string
    - last_name: string
    - email: string
    - phone_number: string
    - about_project: string
    
    Redirects to appropriate calendar URL based on score (1, 2, or 3)
    """
    try:
        # Get data from query parameters
        lead_data = {
            'first_name': request.args.get('first_name', ''),
            'last_name': request.args.get('last_name', ''),
            'email': request.args.get('email', ''),
            'phone_number': request.args.get('phone_number', ''),
            'about_project': request.args.get('about_project', '')
        }
        
        # Validate required fields
        required_fields = ['first_name', 'last_name', 'email', 'phone_number', 'about_project']
        missing_fields = [field for field in required_fields if not lead_data.get(field)]
        
        if missing_fields:
            return f"Missing required query parameters: {', '.join(missing_fields)}", 400
        
        # Qualify the lead
        result = qualify_lead(lead_data)
        
        # Log the request and response
        log_request(lead_data, result)
        
        # Get the score and redirect to appropriate URL with pre-filled data
        score = result.get('score', 0)
        redirect_url = get_calendar_url(score, lead_data)
        
        return redirect(redirect_url)
        
    except Exception as e:
        return f"Internal server error: {str(e)}", 500

@app.route('/qualify-lead', methods=['POST'])
def qualify_lead_endpoint():
    """
    API endpoint to qualify a lead
    
    Expected JSON payload:
    {
        "first_name": "string",
        "last_name": "string", 
        "brand_name": "string",
        "website": "string",
        "email": "string",
        "phone_number": "string",
        "heard_from": "string",
        "designer_type": "string",
        "about_project": "string"
    }
    
    Returns:
    {
        "score": 0, 1, 2, or 3,
        "confidence": "high/medium/low",
        "reason": "string"
    }
    """
    try:
        # Get JSON data from request
        lead_data = request.get_json()
        
        if not lead_data:
            return jsonify({
                "error": "No JSON data provided"
            }), 400
        
        # Validate required fields (simplified)
        required_fields = ['first_name', 'last_name', 'email', 'phone_number', 'about_project']
        missing_fields = [field for field in required_fields if not lead_data.get(field)]
        
        if missing_fields:
            return jsonify({
                "error": f"Missing required fields: {', '.join(missing_fields)}"
            }), 400
        
        # Qualify the lead
        result = qualify_lead(lead_data)
        
        # Log the request and response
        log_request(lead_data, result)
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            "error": f"Internal server error: {str(e)}"
        }), 500

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        "status": "healthy",
        "message": "Lead qualifier API is running"
    })

@app.route('/', methods=['GET'])
def root():
    """Root endpoint with API information"""
    return jsonify({
        "message": "Jennifer's Lead Qualifier API",
        "endpoints": {
            "GET /qualify": "Qualify a lead via query parameters and redirect to calendar (NEW)",
            "POST /qualify-lead": "Qualify a lead and return JSON score (0, 1, 2, or 3)",
            "GET /health": "Health check endpoint",
            "GET /": "This information endpoint"
        },
        "required_fields": [
            "first_name", "last_name", "email", "phone_number", "about_project"
        ],
        "scoring": {
            "0": "Spam",
            "1": "Not right fit",
            "2": "Unsure", 
            "3": "Right fit"
        },
        "example_urls": {
            "GET": "/qualify?first_name=John&last_name=Doe&email=john@example.com&phone_number=555-1234&about_project=I need help with luxury fashion production",
            "POST": "Send JSON to /qualify-lead with the required fields"
        }
    })

if __name__ == '__main__':
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=False)
