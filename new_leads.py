import streamlit as st
import json
import os
import openai

# Initialize the OpenAI client with the API key from environment variables.
client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# New, comprehensive prompt template
PROMPT_TEMPLATE = """
ANALYZE THE FOLLOWING LEAD INFORMATION FROM JENNIFER'S FASHION MANUFACTURING FORM.
CLASSIFY THE LEAD AS EITHER:
SCORE 1: UNQUALIFIED/SPAM - OR - SCORE 3: QUALIFIED/BEST FIT

Lead Information:
- First Name: {first_name}
- Last Name: {last_name}
- Brand Name: {brand_name}
- Website: {website}
- Email: {email}
- Phone Number: {phone_number}
- Heard From: {heard_from}
- Designer Type: {designer_type}
- About Project: {about_project}

CRITERIA FOR SCORE 1 (UNQUALIFIED):
- **Primary Reason:** The lead's project is not a right fit for **The Evans Group (TEG)**, a high-end, luxury fashion manufacturer specializing in couture and small-to-medium volume production. This is the most important factor, even if the request is professional and detailed.
- **Specific examples of "Not Right Fit" projects:**
    - Leads requesting large, mass-market quantities (e.g., thousands of units per style) that do not align with TEG's specialized **small-to-medium volume production model (1-300 pieces per style)**.
    - Projects for basic apparel, sports uniforms, simple activewear, or workwear.
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

OUTPUT FORMAT:
Provide ONLY a JSON response with the following structure, AND NOTHING ELSE:
{{
  "score": 1 or 3,
  "confidence": "high/medium/low",
  "reason": "brief explanation of classification decision"
}}

Focus on the substance, specificity, and professional context of the entire lead profile, and its alignment with TEG's specialized business model.
"""

st.title("Jennifer's Comprehensive Lead Qualifier")
st.markdown("Enter the lead information from the form to get a qualification score.")

# Collect all form fields from the user
first_name = st.text_input("First Name *")
last_name = st.text_input("Last Name *")
brand_name = st.text_input("Brand Name (if you have one)")
website = st.text_input("Website (if you have one)")
email = st.text_input("Email Address *")
phone_number = st.text_input("Phone Number *")

# Create a radio button for "How did you hear about us?"
st.subheader("How did you hear about us? *")
heard_from = st.radio(
    "Select an option:",
    ("None Selected", "Google", "Social", "Textile Show", "Apparel News", "Friend", "CFDA", "Other"),
    index=0
)

# Create a radio button for "What kind of designer are you?"
st.subheader("What kind of designer are you? *")
designer_type = st.radio(
    "Select an option:",
    ("None Selected", "Emerging Designer", "Established High-End Designer", "Somewhere in the Middle", "Other"),
    index=0
)

about_project = st.text_area("Please tell us about your project *", height=200)

if st.button("Classify Lead"):
    # Check for all required fields
    required_fields = {
        "First Name": first_name,
        "Last Name": last_name,
        "Email Address": email,
        "Phone Number": phone_number,
        "How did you hear about us?": heard_from,
        "What kind of designer are you?": designer_type,
        "About Project": about_project
    }

    missing_fields = [field for field, value in required_fields.items() if not value or value == "None Selected"]

    if missing_fields:
        st.warning(f"The following fields are required: {', '.join(missing_fields)}")
    else:
        # Construct the input string for the LLM from all fields
        llm_input = {
            "first_name": first_name,
            "last_name": last_name,
            "brand_name": brand_name,
            "website": website,
            "email": email,
            "phone_number": phone_number,
            "heard_from": heard_from,
            "designer_type": designer_type,
            "about_project": about_project
        }

        # Format the prompt with the collected data
        prompt = PROMPT_TEMPLATE.format(
            first_name=llm_input['first_name'],
            last_name=llm_input['last_name'],
            brand_name=llm_input['brand_name'],
            website=llm_input['website'],
            email=llm_input['email'],
            phone_number=llm_input['phone_number'],
            heard_from=llm_input['heard_from'],
            designer_type=llm_input['designer_type'],
            about_project=llm_input['about_project']
        )

        try:
            # Send the comprehensive prompt to the LLM
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"}
            )

            # Parse the JSON response
            response_json = json.loads(response.choices[0].message.content)

            # Display the results
            st.success("Lead Classification Results:")
            st.json(response_json)

        except json.JSONDecodeError:
            st.error("Failed to parse JSON response from the LLM. Please try again.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
