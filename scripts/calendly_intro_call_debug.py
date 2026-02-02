"""
Diagnostic script for TEG Introductory Call: list all Calendly event types
returned by the API so we can see if Anthony, Heather, Ian appear (by URL or name).
Run from project root: python scripts/calendly_intro_call_debug.py
"""
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
os.chdir(PROJECT_ROOT)
sys.path.insert(0, PROJECT_ROOT)


def load_api_key():
    api_key = None
    secrets_path = os.path.join(PROJECT_ROOT, ".streamlit", "secrets.toml")
    if os.path.exists(secrets_path):
        try:
            import toml
            secrets = toml.load(secrets_path)
        except Exception:
            try:
                import tomllib
                with open(secrets_path, "rb") as f:
                    secrets = tomllib.load(f)
            except Exception:
                pass
        else:
            calendly = secrets.get("calendly", {})
            api_key = calendly.get("calendly_api_key") or calendly.get("api_key")
    if not api_key:
        api_key = os.environ.get("CALENDLY_API_KEY")
    return api_key


def infer_intro_person(name, scheduling_url):
    """Return Anthony, Heather, Ian, or '' based on event type name/URL."""
    name_lower = (name or "").lower()
    url_lower = (scheduling_url or "").lower()
    if "anthony-the-evans-group" in url_lower or "anthony" in name_lower:
        return "Anthony"
    if "heather-the-evans-group" in url_lower or "heather" in name_lower:
        return "Heather"
    if "ian-the-evans-group" in url_lower or ("ian" in name_lower and "christian" not in name_lower):
        return "Ian"
    return ""


def main():
    api_key = load_api_key()
    if not api_key:
        print("No Calendly API key. Set CALENDLY_API_KEY or add to .streamlit/secrets.toml")
        return 1

    import requests
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    print("=" * 60)
    print("CALENDLY EVENT TYPES (for TEG Introductory Call by person)")
    print("=" * 60)

    r = requests.get("https://api.calendly.com/users/me", headers=headers, timeout=30)
    if r.status_code != 200:
        print(f"API users/me failed: {r.status_code} - {r.text[:200]}")
        return 1
    resource = r.json().get("resource", {})
    user_uri = resource.get("uri")
    user_name = resource.get("name", "")
    org_uri = resource.get("current_organization")
    print(f"User: {user_name} ({user_uri})")
    print(f"Organization: {org_uri or '(none)'}\n")

    # Fetch event types: user + organization (so we see Anthony, Heather, Ian if in org)
    event_types = []
    seen = set()
    for param_name, param_value in [("user", user_uri), ("organization", org_uri)]:
        if not param_value:
            continue
        r2 = requests.get(
            "https://api.calendly.com/event_types",
            headers=headers,
            params={param_name: param_value, "count": 100},
            timeout=30,
        )
        if r2.status_code != 200:
            print(f"API event_types?{param_name}=... failed: {r2.status_code}")
            continue
        for et in r2.json().get("collection", []):
            u = et.get("uri")
            if u and u not in seen:
                seen.add(u)
                event_types.append(et)
    print(f"Total event types: {len(event_types)}\n")

    # Intro-call related
    INTRO_MARKERS = ["teg-introductory-call", "introductory", "intro call"]
    print("Event types that look like TEG Introductory Call (name / scheduling_url -> person):")
    print("-" * 60)
    for et in event_types:
        name = et.get("name", "")
        url = et.get("scheduling_url") or ""
        url_lower = url.lower()
        name_lower = name.lower()
        is_intro = any(m in url_lower or m in name_lower for m in INTRO_MARKERS)
        person = infer_intro_person(name, url)
        if is_intro or person:
            tag = f" -> {person}" if person else " -> (no person from URL/name)"
            print(f"  name: {repr(name)}")
            print(f"  url:  {repr(url)}{tag}")
            print()
    # Admin token test: try scheduled_events?organization= (org-wide events; admin tokens grant this)
    if org_uri:
        min_t = "2025-01-01T00:00:00.000000Z"
        max_t = "2026-12-31T23:59:59.999999Z"
        r_org = requests.get("https://api.calendly.com/scheduled_events", headers=headers,
                             params={"organization": org_uri, "min_start_time": min_t, "max_start_time": max_t, "count": 5}, timeout=30)
        if r_org.status_code == 200:
            n = len(r_org.json().get("collection", []))
            print("scheduled_events?organization= (admin scope): OK - got", n, "events (first page)")
        else:
            print("scheduled_events?organization= (admin scope):", r_org.status_code, "-", r_org.text[:150])
    print("\nAll event types (name, slug, scheduling_url, profile, GET user for profile.owner):")
    print("-" * 60)
    seen_owners = {}
    for et in event_types:
        name = et.get("name")
        slug = et.get("slug")
        url = et.get("scheduling_url")
        profile = et.get("profile") or {}
        profile_name = profile.get("name")
        profile_owner = profile.get("owner")
        profile_type = profile.get("type")
        print(f"  name: {repr(name)}  slug: {repr(slug)}")
        print(f"    scheduling_url: {repr(url)}")
        print(f"    profile.type: {repr(profile_type)}  profile.name: {repr(profile_name)}  profile.owner: {repr(profile_owner)}")
        if profile_owner and profile_owner not in seen_owners:
            uuid = profile_owner.rstrip("/").split("/")[-1]
            r = requests.get(f"https://api.calendly.com/users/{uuid}", headers=headers, timeout=10)
            if r.status_code == 200:
                user_resource = r.json().get("resource", {})
                seen_owners[profile_owner] = user_resource
                print(f"    GET user: name={repr(user_resource.get('name'))} slug={repr(user_resource.get('slug'))}")
            else:
                print(f"    GET user: {r.status_code}")
        elif profile_owner and profile_owner in seen_owners:
            u = seen_owners[profile_owner]
            print(f"    GET user (cached): name={repr(u.get('name'))} slug={repr(u.get('slug'))}")
    print("\nDone. Use this to confirm Anthony/Heather/Ian appear so intro call can break down by person.")


if __name__ == "__main__":
    sys.exit(main() or 0)
