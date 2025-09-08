import json
import random
from faker import Faker

fake = Faker()

def make_spam():
    return {
        "first_name": fake.first_name(),
        "last_name": fake.last_name(),
        "brand_name": fake.company(),
        "website": fake.url(),
        "email": fake.free_email(),
        "phone_number": fake.phone_number(),
        "heard_from": random.choice(["Google", "Other", "Unknown"]),
        "designer_type": random.choice(["Other", "None"]),
        "about_project": random.choice([
            "We provide SEO services, please contact us.",
            "Do you need an electrician for your factory?",
            "Hello, we sell cheap bulk T-shirts, 10,000 units minimum.",
            "Offering marketing packages for your business.",
            "Altario_sports_wear is a professional supplier of custom clothing with free branding."
        ]),
        "expected_output": {
            "score": 1,
            "confidence": "high",
            "reason": "Spam/unrelated or bulk low-end apparel not aligned with TEG's luxury small-batch model."
        }
    }

def make_not_right_fit():
    return {
        "first_name": fake.first_name(),
        "last_name": fake.last_name(),
        "brand_name": fake.company(),
        "website": fake.url(),
        "email": fake.company_email(),
        "phone_number": fake.phone_number(),
        "heard_from": random.choice(["Instagram", "Facebook", "Google"]),
        "designer_type": random.choice(["Sportswear", "Workwear", "Medical"]),
        "about_project": random.choice([
            "We need 5,000 gym uniforms produced every month.",
            "Looking for mass production of 20,000 T-shirts.",
            "We design hospital scrubs and need a manufacturer.",
            "I want simple polo shirts for a workwear company.",
            "Supplier needed for bulk underwear manufacturing."
        ]),
        "expected_output": {
            "score": 1,
            "confidence": "high",
            "reason": "Legitimate inquiry but outside TEG’s scope (mass-market, uniforms, medical/workwear)."
        }
    }

def make_fit():
    return {
        "first_name": fake.first_name(),
        "last_name": fake.last_name(),
        "brand_name": fake.company(),
        "website": fake.url(),
        "email": fake.company_email(),
        "phone_number": fake.phone_number(),
        "heard_from": random.choice(["Instagram", "Google", "Referral", "Trade Show"]),
        "designer_type": random.choice(["Luxury", "Emerging Designer", "High-end Couture"]),
        "about_project": random.choice([
            "I am launching a bridal line with silk gowns, looking for 30 custom pieces.",
            "We need help producing a capsule collection of 15 runway looks.",
            "Seeking high-end evening wear production with lace and beading.",
            "Launching a menswear line with 50 tailored jackets and trousers.",
            "Creating a luxury womenswear brand with delicate silks and wovens."
        ]),
        "expected_output": {
            "score": 3,
            "confidence": "high",
            "reason": "Detailed project aligned with TEG’s couture/small-batch luxury production model."
        }
    }

# Generate dataset
data = []
for _ in range(200):
    data.append(make_spam())
for _ in range(200):
    data.append(make_not_right_fit())
for _ in range(200):
    data.append(make_fit())

# Shuffle the dataset so categories are mixed
random.shuffle(data)

# Save to JSON file
with open("test_qualifier.json", "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)

print("✅ test_qualifier.json with 600 annotated leads generated!")
