from flask import Flask, request, jsonify, redirect
import json
import openai
import toml
import os
import csv
from datetime import datetime

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

# Qualification prompt template (simplified version)
PROMPT_TEMPLATE = """
ANALYZE THE FOLLOWING LEAD INFORMATION FROM JENNIFER'S FASHION MANUFACTURING FORM.
CLASSIFY THE LEAD AS:
SCORE 1: UNQUALIFIED/SPAM
SCORE 2: UNSURE (not clearly spam but not clearly qualified either)
SCORE 3: QUALIFIED/BEST FIT

Lead Information:
- First Name: {first_name}
- Last Name: {last_name}
- Email: {email}
- Phone Number: {phone_number}
- About Project: {about_project}

CRITERIA FOR SCORE 1 (UNQUALIFIED):
- **Primary Reason:** The lead's project is not a right fit for **The Evans Group (TEG)**, a high-end, luxury fashion manufacturer specializing in couture and small-to-medium volume production. This is the most important factor, even if the request is professional and detailed.
- **Specific examples of "Not Right Fit" projects:**
    - Leads requesting large, mass-market quantities (e.g., thousands of units per style) that do not align with TEG's specialized **small-to-medium volume production model (1-300 pieces per style)**.
    - Projects for basic apparel, sports, gymnastics and/or rhythmic uniforms, simple activewear, or workwear.
    - Medical apparel, underwear, or other non-fashion garments.
    - General spam, such as advertising services, or unrelated inquiries (e.g., "Do you need an electrician?").
    - Vague or generic project descriptions that lack substance, details, or a clear vision, regardless of the lead's professional title.

- **Unprofessional or missing information:** A generic, non-existent, or irrelevant brand name and website. Generic email addresses (e.g., disposable mail services) or fake phone numbers.

CRITERIA FOR SCORE 3 (QUALIFIED):
- **Primary Reason:** The "About Project" description is detailed and aligns with TEG's services for **high-end, luxury, and couture fashion brands**. This is a crucial requirement.
- **Project alignment:** The lead's request fits with TEG's services, which include:
    - **Client Types:** Established high-end designers, luxury brands, or emerging designers with a professional vision.
    - **Production Scale:** The request aligns with **small-to-medium volume production (1-300 pieces per style)**, or other TEG services like pattern making or sample creation.
    - **Specific products:** Runway looks, bridal, evening wear, formal wear, high-end menswear and womenswear, complex knitwear, or specific, elevated basics.
    - **Specific materials:** Delicate silks, leather, wovens, beading, lace, or complex knits.
    - **Project Details:** The request includes specific details on materials, quantities, design elements, and timelines.
- **Professionalism:** The lead provides clear, professional information across all fields (brand name, website, and designer type).

CRITERIA FOR SCORE 2 (UNSURE):
- **When to use:** The lead is clearly not spam or completely unfit, but there's not enough information to confidently classify as qualified.
- **Examples:** Vague project descriptions that could be legitimate but lack detail, or projects that might align with TEG but need more information.

CRITERIA FOR SCORE 3 (QUALIFIED):
- **Primary Reason:** The "About Project" description is detailed and aligns with TEG's services for **high-end, luxury, and couture fashion brands**. This is a crucial requirement.
- **Project alignment:** The lead's request fits with TEG's services, which include:
    - **Client Types:** Established high-end designers, luxury brands, or emerging designers with a professional vision.
    - **Production Scale:** The request aligns with **small-to-medium volume production (1-300 pieces per style)**, or other TEG services like pattern making or sample creation.
    - **Specific products:** Runway looks, bridal, evening wear, formal wear, high-end menswear and womenswear, complex knitwear, or specific, elevated basics.
    - **Specific materials:** Delicate silks, leather, wovens, beading, lace, or complex knits.

OUTPUT FORMAT:
Provide ONLY a JSON response with the following structure, AND NOTHING ELSE:
{{
  "score": 1, 2, or 3,
  "confidence": "high/medium/low",
  "reason": "brief explanation of classification decision"
}}

Focus on the substance, specificity, and professional context of the lead profile, and its alignment with TEG's specialized business model.
"""

# Calendar URLs for redirection based on score
CALENDAR_URLS = {
    1: "https://www.google.com",  # Dummy URL for score 1 (unqualified)
    2: "https://www.yahoo.com",   # Dummy URL for score 2 (unsure)
    3: "https://www.gmail.com"    # Dummy URL for score 3 (qualified)
}

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
                "score": 1,
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
            "score": 1,
            "confidence": "low",
            "reason": f"Error parsing AI response: {str(e)}"
        }
    except Exception as e:
        return {
            "score": 1,
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
        
        # Get the score and redirect to appropriate URL
        score = result.get('score', 1)
        redirect_url = CALENDAR_URLS.get(score, CALENDAR_URLS[1])  # Default to score 1 URL
        
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
        "score": 1 or 3,
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
            "POST /qualify-lead": "Qualify a lead and return JSON score (1, 2, or 3)",
            "GET /health": "Health check endpoint",
            "GET /": "This information endpoint"
        },
        "required_fields": [
            "first_name", "last_name", "email", "phone_number", "about_project"
        ],
        "scoring": {
            "1": "Unqualified/Spam - redirects to Google",
            "2": "Unsure - redirects to Yahoo", 
            "3": "Qualified/Best Fit - redirects to Gmail"
        },
        "example_urls": {
            "GET": "/qualify?first_name=John&last_name=Doe&email=john@example.com&phone_number=555-1234&about_project=I need help with luxury fashion production",
            "POST": "Send JSON to /qualify-lead with the required fields"
        }
    })

if __name__ == '__main__':
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=False)
