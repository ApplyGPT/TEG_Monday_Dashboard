import streamlit as st
import json
import os
import openai
import pandas as pd
import time

# Set page config at the top of the script
st.set_page_config(page_title="Jennifer's Lead Qualifier", layout="wide")

# Load the API key from Streamlit secrets
try:
    openai_api_key = st.secrets["openai"]["api_key"]
    openai.api_key = openai_api_key
except KeyError:
    st.error("OpenAI API key not found in secrets.toml. Please ensure the [openai] section with api_key is configured.")
    st.stop()
except Exception as e:
    st.error(f"An error occurred while loading the API key from secrets: {e}")
    st.stop()

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
st.subheader("What kind of designer are you?")
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
            response = openai.chat.completions.create(
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

# --- Section 2: Batch test all JSON files ---
st.header("🔹 Run Batch Test on JSON Files")

test_folder = st.text_input("Enter test folder path (e.g., test_leads/)", "test_leads/")

if st.button("Run Batch Test"):
    if not os.path.exists(test_folder):
        st.error(f"Folder not found: {test_folder}")
    else:
        results = []
        # Sort filenames numerically (works for file_1.json ... file_10.json)
        json_files = sorted(
            [f for f in os.listdir(test_folder) if f.endswith(".json")],
            key=lambda x: int("".join(filter(str.isdigit, x)) or 0)
        )

        for file_name in json_files:
            file_path = os.path.join(test_folder, file_name)
            with open(file_path, "r", encoding="utf-8") as f:
                lead_data = json.load(f)

            prompt = PROMPT_TEMPLATE.format(
                first_name=lead_data.get("first_name", ""),
                last_name=lead_data.get("last_name", ""),
                brand_name=lead_data.get("brand_name", ""),
                website=lead_data.get("website", ""),
                email=lead_data.get("email", ""),
                phone_number=lead_data.get("phone_number", ""),
                heard_from=lead_data.get("heard_from", ""),
                designer_type=lead_data.get("designer_type", ""),
                about_project=lead_data.get("about_project", "")
            )

            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}],
                    response_format={"type": "json_object"}
                )
                response_json = json.loads(response.choices[0].message.content)
                results.append({"file": file_name, **response_json})
            except Exception as e:
                results.append({"file": file_name, "error": str(e)})

        st.success("Batch Results:")
        st.json(results)

# --- Section 3: Batch test on a single JSON file with accuracy evaluation ---
st.header("🔹 Run Batch Test on a Single JSON File (with Accuracy Metrics)")

if st.button("Run Test on test_qualifier.json"):
    test_file_path = "test_qualifier.json"
    
    if not os.path.exists(test_file_path):
        st.error(f"File not found: {test_file_path}. Please ensure the file is in the same directory as the app.")
    else:
        try:
            with open(test_file_path, "r", encoding="utf-8") as f:
                leads_data = json.load(f)

            results = []
            total_rows = len(leads_data)
            progress_bar = st.progress(0)
            status_text = st.empty()

            correct = 0
            total = 0
            # Counters for per-category accuracy
            category_stats = {
                "spam": {"correct": 0, "total": 0},
                "not_right_fit": {"correct": 0, "total": 0},
                "fit": {"correct": 0, "total": 0}
            }

            for idx, lead_data in enumerate(leads_data):
                status_text.text(f"Processing lead {idx + 1} of {total_rows}...")

                # Ground truth (expected)
                expected = lead_data.get("expected_output", {})
                expected_score = expected.get("score")

                # Build prompt
                prompt = PROMPT_TEMPLATE.format(
                    first_name=lead_data.get("first_name", ""),
                    last_name=lead_data.get("last_name", ""),
                    brand_name=lead_data.get("brand_name", ""),
                    website=lead_data.get("website", ""),
                    email=lead_data.get("email", ""),
                    phone_number=lead_data.get("phone_number", ""),
                    heard_from=lead_data.get("heard_from", ""),
                    designer_type=lead_data.get("designer_type", ""),
                    about_project=lead_data.get("about_project", "")
                )

                try:
                    response = openai.chat.completions.create(
                        model="gpt-4.1-mini",
                        messages=[{"role": "user", "content": prompt}],
                        response_format={"type": "json_object"}
                    )
                    response_json = json.loads(response.choices[0].message.content)
                    predicted_score = response_json.get("score")

                    # Track accuracy
                    total += 1
                    if predicted_score == expected_score:
                        correct += 1

                        # Category-specific check
                        if expected_score == 1 and "spam" in expected.get("reason", "").lower():
                            category_stats["spam"]["correct"] += 1
                        elif expected_score == 1:
                            category_stats["not_right_fit"]["correct"] += 1
                        elif expected_score == 3:
                            category_stats["fit"]["correct"] += 1

                    # Update totals per category
                    if expected_score == 1 and "spam" in expected.get("reason", "").lower():
                        category_stats["spam"]["total"] += 1
                    elif expected_score == 1:
                        category_stats["not_right_fit"]["total"] += 1
                    elif expected_score == 3:
                        category_stats["fit"]["total"] += 1

                    results.append({
                        "predicted": response_json,
                        "expected": expected,
                        "original_data": lead_data
                    })

                except Exception as e:
                    results.append({
                        "predicted": {"score": None, "reason": f"Error: {e}"},
                        "expected": expected,
                        "original_data": lead_data
                    })

                progress_bar.progress((idx + 1) / total_rows)
                time.sleep(0.01)

            progress_bar.empty()
            status_text.empty()

            # Overall accuracy
            overall_acc = correct / total if total > 0 else 0.0

            # Category accuracies
            spam_acc = category_stats["spam"]["correct"] / category_stats["spam"]["total"] if category_stats["spam"]["total"] > 0 else 0
            nrf_acc = category_stats["not_right_fit"]["correct"] / category_stats["not_right_fit"]["total"] if category_stats["not_right_fit"]["total"] > 0 else 0
            fit_acc = category_stats["fit"]["correct"] / category_stats["fit"]["total"] if category_stats["fit"]["total"] > 0 else 0

            st.success(f"✅ Batch classification completed! Processed {total_rows} leads.")
            st.write("### 📊 Accuracy Metrics")
            st.write(f"**Overall Accuracy:** {overall_acc:.2%}")
            st.write(f"**Spam Accuracy:** {spam_acc:.2%} ({category_stats['spam']['correct']}/{category_stats['spam']['total']})")
            st.write(f"**Not Right Fit Accuracy:** {nrf_acc:.2%} ({category_stats['not_right_fit']['correct']}/{category_stats['not_right_fit']['total']})")
            st.write(f"**Fit Accuracy:** {fit_acc:.2%} ({category_stats['fit']['correct']}/{category_stats['fit']['total']})")

            # Display first few results as sample
            st.write("### Sample Results")
            st.json(results[:10])

            # Download all results
            json_out = json.dumps(results, indent=2, ensure_ascii=False).encode("utf-8")
            st.download_button(
                "Download Classified JSON (with predictions + expected)",
                json_out,
                "classified_leads_with_expected.json",
                "application/json"
            )

        except json.JSONDecodeError:
            st.error("Failed to parse the JSON file. Please ensure it's in a valid JSON format.")
        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
