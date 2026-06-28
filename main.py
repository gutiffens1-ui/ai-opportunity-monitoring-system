import argparse
import base64
import csv
import importlib.util
import json
import os
import re
from datetime import date
from email.mime.text import MIMEText
from pathlib import Path
from urllib.parse import quote_plus

try:
    import requests
except ImportError:
    requests = None

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None


try:
    import feedparser
except ImportError:
    feedparser = None


try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except ImportError:
    Request = None
    Credentials = None
    InstalledAppFlow = None
    build = None


BASE_DIR = Path(__file__).resolve().parent

OPPORTUNITY_INPUT_FIELDS = [
    "company",
    "website",
    "description",
    "contact_name",
    "contact_title",
    "contact_link",
    "recipient_email",
]

OPPORTUNITY_OUTPUT_FIELDS = OPPORTUNITY_INPUT_FIELDS + [
    "relevance_score",
    "opportunity_type",
    "why_relevant",
    "suggested_outreach_angle",
    "linkedin_connect_note",
    "cold_email_subject",
    "cold_email_body",
    "ready_for_gmail",
]

SEED_COMPANY_FIELDS = [
    "company",
    "company_website",
    "description",
    "target_contact_titles",
    "suggested_contacts",
    "public_search_links",
    "rss_feed_url",
    "notes",
]

DAILY_TARGET_FIELDS = [
    "contact_name",
    "contact_title",
    "company",
    "company_website",
    "contact_link",
    "public_email",
    "why_relevant",
    "confidence_score",
    "suggested_outreach_angle",
    "linkedin_connect_note",
    "cold_email_subject",
    "cold_email_body",
    "ready_for_gmail",
    "sourced_from",
    "source_notes",
    "date_added",
]

LINKEDIN_SEARCH_FIELDS = [
    "company",
    "company_website",
    "target_title",
    "linkedin_search_query",
    "linkedin_search_url",
    "why_this_search_matters",
    "manual_review_status",
    "notes",
]

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

REQUIRED_CSV_HEADERS = {
    "opportunities.csv": OPPORTUNITY_INPUT_FIELDS,
    "seed_companies.csv": SEED_COMPANY_FIELDS,
    "output.csv": OPPORTUNITY_OUTPUT_FIELDS,
    "daily_targets.csv": DAILY_TARGET_FIELDS,
    "contact_database.csv": DAILY_TARGET_FIELDS,
    "linkedin_search_tasks.csv": LINKEDIN_SEARCH_FIELDS,
}

REQUIRED_PACKAGES = {
    "openai": "openai",
    "python-dotenv": "dotenv",
    "requests": "requests",
    "beautifulsoup4": "bs4",
    "feedparser": "feedparser",
    "google-api-python-client": "googleapiclient",
    "google-auth-httplib2": "google_auth_httplib2",
    "google-auth-oauthlib": "google_auth_oauthlib",
}

SAMPLE_MARKERS = ("SAMPLE", "EXAMPLE", "REPLACE WITH")

PRIORITY_TITLES = [
    "Head of Production",
    "Producer",
    "Executive Producer",
    "Creative Producer",
    "Development Producer",
    "Production Manager",
    "Head of Content",
    "VP of Content",
    "Content Acquisition Manager",
    "Partnerships Manager",
    "Business Development Manager",
    "Founder",
    "CEO",
]

LINKEDIN_TARGET_TITLES = [
    "Producer",
    "Executive Producer",
    "Creative Producer",
    "Development Producer",
    "Head of Production",
    "Production Manager",
    "Head of Content",
    "VP of Content",
    "Content Acquisition Manager",
    "Partnerships Manager",
    "Business Development Manager",
    "Founder",
    "CEO",
]


def read_csv_rows(filename):
    path = BASE_DIR / filename
    if not path.exists():
        return []

    with path.open("r", newline="", encoding="utf-8-sig") as file:
        return list(csv.DictReader(file))


def write_csv_rows(filename, fieldnames, rows):
    path = BASE_DIR / filename
    # UTF-8 with a BOM helps Excel on Windows recognize text cleanly.
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def append_unique_contacts(filename, fieldnames, new_rows):
    existing_rows = [row for row in read_csv_rows(filename) if not is_sample_row(row)]
    seen = {contact_key(row) for row in existing_rows}
    rows_to_add = []

    for row in new_rows:
        if is_sample_row(row):
            continue
        key = contact_key(row)
        if key not in seen:
            rows_to_add.append(row)
            seen.add(key)

    write_csv_rows(filename, fieldnames, existing_rows + rows_to_add)
    return len(rows_to_add)


def contact_key(row):
    return (
        clean_key(row.get("contact_name", "")),
        clean_key(row.get("company", "")),
        clean_key(row.get("contact_link", "")),
    )


def is_sample_row(row):
    text = " ".join(str(value or "") for value in row.values()).upper()
    return any(marker in text for marker in SAMPLE_MARKERS)


def clean_key(value):
    return " ".join(str(value or "").lower().strip().split())


def get_openai_client():
    if OpenAI is None:
        return None
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


def ask_openai_for_json(system_message, user_message, fallback, quiet=False):
    client = get_openai_client()
    if client is None:
        if not quiet:
            print("OPENAI_API_KEY or the OpenAI package is missing. I saved beginner-friendly placeholder results instead.")
        return fallback

    model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    try:
        response = client.chat.completions.create(
            model=model,
            temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": user_message},
            ],
        )
        content = response.choices[0].message.content
        return json.loads(content)
    except Exception as error:
        if not quiet:
            print(f"OpenAI analysis failed, so I used a safe placeholder. Error: {error}")
        return fallback


def analyze_opportunities():
    rows = read_csv_rows("opportunities.csv")
    analyzed_rows = []

    for row in rows:
        if is_sample_row(row):
            continue

        fallback = {
            "relevance_score": "5",
            "opportunity_type": "Needs OpenAI review",
            "why_relevant": "Add OPENAI_API_KEY to .env, then run the analysis again.",
            "suggested_outreach_angle": "Ask whether they are exploring short-form vertical drama or AI-assisted production.",
            "linkedin_connect_note": make_connect_note(row.get("company", ""), row.get("contact_name", "")),
            "cold_email_subject": f"Short-form vertical drama idea for {row.get('company', 'your team')}",
            "cold_email_body": make_basic_email(row),
        }

        prompt = {
            "task": "Analyze this opportunity for short-form vertical drama production, AI short drama projects, content acquisition, production partnerships, or business development.",
            "row": row,
            "required_output": {
                "relevance_score": "integer from 1 to 10",
                "opportunity_type": "short label",
                "why_relevant": "2 to 4 plain English sentences",
                "suggested_outreach_angle": "specific outreach angle",
                "linkedin_connect_note": "under 280 characters",
                "cold_email_subject": "short subject line",
                "cold_email_body": "brief email draft, human-reviewed before sending",
            },
        }

        result = ask_openai_for_json(
            "You are a careful business development research assistant. Return only valid JSON. Do not invent private information.",
            json.dumps(prompt, indent=2),
            fallback,
        )

        combined = dict(row)
        for field in OPPORTUNITY_OUTPUT_FIELDS:
            if field not in combined:
                combined[field] = result.get(field, "")

        final_outreach = run_outreach_quality_loop(
            row,
            combined.get("linkedin_connect_note", ""),
            combined.get("cold_email_subject", ""),
            combined.get("cold_email_body", ""),
        )
        combined.update(final_outreach)
        combined["ready_for_gmail"] = "FALSE"
        analyzed_rows.append(combined)

    write_csv_rows("output.csv", OPPORTUNITY_OUTPUT_FIELDS, analyzed_rows)
    print(f"Saved {len(analyzed_rows)} analyzed opportunities to output.csv.")


def source_daily_contacts():
    seed_rows = read_csv_rows("seed_companies.csv")
    all_candidates = []
    notes = []

    for seed in seed_rows:
        if is_sample_row(seed):
            continue

        evidence = collect_public_evidence(seed)
        suggested_contacts = parse_suggested_contacts(seed)

        prompt = {
            "task": "Find up to 10 high-quality contacts from the evidence below. Only include named people when their name appears in the evidence or in suggested_contacts. Do not scrape LinkedIn. Do not invent emails.",
            "priority_titles": PRIORITY_TITLES,
            "seed_company": seed,
            "suggested_contacts": suggested_contacts,
            "public_evidence": evidence,
            "required_contact_fields": DAILY_TARGET_FIELDS,
        }

        fallback_contacts = build_fallback_contacts(seed, suggested_contacts)
        result = ask_openai_for_json(
            "You are a careful contact sourcing assistant. Return only JSON with a contacts array. Use only public or user-provided information.",
            json.dumps(prompt, indent=2),
            {"contacts": fallback_contacts},
        )

        contacts = result.get("contacts", [])
        if not isinstance(contacts, list):
            contacts = []

        for contact in contacts:
            clean_contact = normalize_contact(contact, seed)
            if not is_sample_row(clean_contact) and is_good_contact(clean_contact):
                all_candidates.append(clean_contact)

    unique_candidates = unique_contacts(all_candidates)
    daily_targets = unique_candidates[:10]

    for contact in daily_targets:
        final_outreach = run_outreach_quality_loop(
            contact,
            contact.get("linkedin_connect_note", ""),
            contact.get("cold_email_subject", ""),
            contact.get("cold_email_body", ""),
        )
        contact.update(final_outreach)
        contact["ready_for_gmail"] = "FALSE"

    write_csv_rows("daily_targets.csv", DAILY_TARGET_FIELDS, daily_targets)
    added = append_unique_contacts("contact_database.csv", DAILY_TARGET_FIELDS, daily_targets)

    if len(daily_targets) < 10:
        notes.append(
            f"Only {len(daily_targets)} good contacts were saved because the available public/user-provided evidence did not support 10 confident contacts."
        )
    if not daily_targets:
        notes.append(
            "No contacts were found from the available public evidence. Run: py main.py linkedin-search"
        )
        notes.append(
            "Review the LinkedIn search links manually, then add verified contacts back into seed_companies.csv."
        )

    print(f"Saved {len(daily_targets)} contacts to daily_targets.csv.")
    print(f"Added {added} new unique contacts to contact_database.csv.")
    for note in notes:
        print(note)


def generate_linkedin_search_tasks():
    """Create manual LinkedIn People search links without visiting LinkedIn."""
    seed_rows = read_csv_rows("seed_companies.csv")
    existing_tasks = {
        (
            clean_key(row.get("company", "")),
            clean_key(row.get("target_title", "")),
        ): row
        for row in read_csv_rows("linkedin_search_tasks.csv")
        if not is_sample_row(row)
    }
    tasks = []

    for seed in seed_rows:
        if is_sample_row(seed):
            continue

        company = seed.get("company", "").strip()
        if not company:
            continue

        website = seed.get("company_website", "").strip()
        for target_title in LINKEDIN_TARGET_TITLES:
            query = f'"{company}" "{target_title}"'
            previous = existing_tasks.get(
                (clean_key(company), clean_key(target_title)), {}
            )
            tasks.append(
                {
                    "company": company,
                    "company_website": website,
                    "target_title": target_title,
                    "linkedin_search_query": query,
                    "linkedin_search_url": (
                        "https://www.linkedin.com/search/results/people/"
                        f"?keywords={quote_plus(query)}"
                    ),
                    "why_this_search_matters": linkedin_search_reason(
                        company, target_title
                    ),
                    "manual_review_status": previous.get("manual_review_status")
                    or "Not reviewed",
                    "notes": previous.get("notes", ""),
                }
            )

    write_csv_rows("linkedin_search_tasks.csv", LINKEDIN_SEARCH_FIELDS, tasks)
    print(f"Saved {len(tasks)} manual LinkedIn search tasks to linkedin_search_tasks.csv.")
    print("No LinkedIn pages were opened or scraped. Review each search manually in your browser.")


def linkedin_search_reason(company, target_title):
    title = target_title.lower()
    if "producer" in title or "production" in title:
        focus = "production planning, creative development, or production partnerships"
    elif "content" in title:
        focus = "content strategy, acquisition, or short-form programming decisions"
    else:
        focus = "company partnerships, business development, or senior decisions"
    article = "An" if target_title[:1].lower() in "aeiou" else "A"
    return f"{article} {target_title} at {company} may influence {focus}."


def collect_public_evidence(seed):
    evidence = []
    evidence.append(
        {
            "source": "seed_companies.csv",
            "text": f"{seed.get('company', '')}\n{seed.get('description', '')}\n{seed.get('notes', '')}",
        }
    )

    if os.getenv("FETCH_WEBSITES", "false").lower() == "true":
        website = seed.get("company_website", "")
        if website:
            evidence.append(fetch_public_page_text(website, "company website"))

    if os.getenv("FETCH_PUBLIC_LINKS", "false").lower() == "true":
        for link in split_multi_value(seed.get("public_search_links", "")):
            evidence.append(fetch_public_page_text(link, "public link"))

    rss_url = seed.get("rss_feed_url", "").strip()
    if rss_url and feedparser is not None:
        parsed = feedparser.parse(rss_url)
        entries = []
        for entry in parsed.entries[:5]:
            entries.append(f"{entry.get('title', '')}: {entry.get('summary', '')}")
        evidence.append({"source": rss_url, "text": "\n".join(entries)})

    serp_api_key = os.getenv("SERPAPI_API_KEY")
    if serp_api_key:
        query = build_search_query(seed)
        evidence.extend(search_with_serpapi(query, serp_api_key))

    return evidence


def fetch_public_page_text(url, source_label):
    if "linkedin.com" in url.lower():
        return {"source": url, "text": "Skipped LinkedIn URL. This project does not scrape LinkedIn."}

    try:
        response = requests.get(url, timeout=10, headers={"User-Agent": "AI Opportunity Radar beginner research tool"})
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        text = " ".join(soup.get_text(" ").split())
        return {"source": f"{source_label}: {url}", "text": text[:4000]}
    except Exception as error:
        return {"source": url, "text": f"Could not read this page: {error}"}


def search_with_serpapi(query, api_key):
    try:
        response = requests.get(
            "https://serpapi.com/search.json",
            params={"q": query, "api_key": api_key, "num": 5},
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
    except Exception as error:
        return [{"source": "SerpAPI", "text": f"Search failed: {error}"}]

    evidence = []
    for item in data.get("organic_results", [])[:5]:
        link = item.get("link", "")
        if "linkedin.com" in link.lower():
            continue
        evidence.append(
            {
                "source": link,
                "text": f"{item.get('title', '')}\n{item.get('snippet', '')}",
            }
        )
    return evidence


def build_search_query(seed):
    company = seed.get("company", "")
    titles = " OR ".join([f'"{title}"' for title in PRIORITY_TITLES[:8]])
    return f'"{company}" ({titles}) short-form vertical drama production content partnerships'


def parse_suggested_contacts(seed):
    contacts = []
    for item in split_multi_value(seed.get("suggested_contacts", "")):
        parts = [part.strip() for part in item.split("|")]
        contacts.append(
            {
                "contact_name": parts[0] if len(parts) > 0 else "",
                "contact_title": parts[1] if len(parts) > 1 else "",
                "contact_link": parts[2] if len(parts) > 2 else "",
                "public_email": parts[3] if len(parts) > 3 else "",
            }
        )
    return contacts


def split_multi_value(value):
    return [item.strip() for item in str(value or "").split(";") if item.strip()]


def build_fallback_contacts(seed, suggested_contacts):
    contacts = []
    for contact in suggested_contacts:
        row = normalize_contact(contact, seed)
        row["why_relevant"] = "This contact was provided in seed_companies.csv and should be reviewed before outreach."
        row["confidence_score"] = "6"
        row["suggested_outreach_angle"] = "Ask whether they are open to short-form vertical drama or AI-assisted production conversations."
        row["linkedin_connect_note"] = make_connect_note(seed.get("company", ""), row.get("contact_name", ""))
        row["cold_email_subject"] = f"Short-form vertical drama idea for {seed.get('company', 'your team')}"
        row["cold_email_body"] = make_basic_email(
            {
                "company": seed.get("company", ""),
                "contact_name": row.get("contact_name", ""),
                "contact_title": row.get("contact_title", ""),
                "description": seed.get("description", ""),
            }
        )
        contacts.append(row)
    return contacts


def normalize_contact(contact, seed):
    row = {field: contact.get(field, "") for field in DAILY_TARGET_FIELDS}
    row["company"] = row.get("company") or seed.get("company", "")
    row["company_website"] = row.get("company_website") or seed.get("company_website", "")
    row["sourced_from"] = row.get("sourced_from") or "seed_companies.csv/public evidence"
    row["date_added"] = row.get("date_added") or date.today().isoformat()
    row["linkedin_connect_note"] = limit_text(row.get("linkedin_connect_note", ""), 279)
    row["ready_for_gmail"] = "FALSE"
    return row


def is_good_contact(contact):
    if not contact.get("contact_name") or contact.get("contact_name", "").lower() in ["unknown", "n/a"]:
        return False
    if not contact.get("company"):
        return False

    score_text = str(contact.get("confidence_score", "0"))
    score_match = re.search(r"\d+", score_text)
    score = int(score_match.group(0)) if score_match else 0
    return score >= get_min_contact_score()


def get_min_contact_score():
    value = os.getenv("MIN_CONTACT_SCORE", "7")
    try:
        score = int(value)
    except ValueError:
        return 7
    return score if 1 <= score <= 10 else 7


def unique_contacts(rows):
    seen = set()
    unique_rows = []
    for row in rows:
        key = contact_key(row)
        if key not in seen:
            seen.add(key)
            unique_rows.append(row)
    return unique_rows


def make_connect_note(company, contact_name):
    greeting = f"Hi {contact_name}," if contact_name else "Hi,"
    return limit_text(
        f"{greeting} I am exploring short-form vertical drama and AI-assisted production partnerships and thought {company} looked relevant. Would be glad to connect.",
        279,
    )


def make_basic_email(row):
    name = row.get("contact_name") or "there"
    company = row.get("company") or "your team"
    description = row.get("description") or "your content work"
    return (
        f"Hi {name},\n\n"
        f"I came across {company} and noticed {description}.\n\n"
        "I am exploring short-form vertical drama and AI-assisted production opportunities, including production partnerships, content acquisition, and business development.\n\n"
        "If this is relevant, would you be open to a short conversation?\n\n"
        "Best,\n"
        "Your Name"
    )


def run_outreach_quality_loop(context, linkedin_note, subject, email_body):
    """Review and revise outreach at most twice, then return only the final draft."""
    current = {
        "linkedin_connect_note": limit_text(linkedin_note, 279),
        "cold_email_subject": str(subject or "").strip(),
        "cold_email_body": str(email_body or "").strip(),
    }

    for _ in range(2):
        local_issues = outreach_quality_issues(
            current["linkedin_connect_note"], current["cold_email_body"]
        )
        safe_revision = make_safe_outreach_revision(context)
        fallback = {
            "passes_quality_check": not local_issues,
            "issues": local_issues,
            **(current if not local_issues else safe_revision),
        }
        review_prompt = {
            "task": "Review this outreach. If any check fails, revise all three draft fields. Use only facts in the context.",
            "context": context,
            "draft": current,
            "checks": [
                "LinkedIn note is under 280 characters",
                "sounds natural",
                "is not too salesy",
                "value proposition is clear",
                "email has exactly one simple call to action",
                "contains no fake or unsupported claims",
                "contains no spammy language",
            ],
            "required_output": {
                "passes_quality_check": "true or false",
                "issues": "list of short issue descriptions",
                "linkedin_connect_note": "final note under 280 characters",
                "cold_email_subject": "short natural subject",
                "cold_email_body": "brief email with one call to action",
            },
        }
        review = ask_openai_for_json(
            "You are a strict outreach quality editor. Return only valid JSON. Never add facts that are not in the supplied context.",
            json.dumps(review_prompt, indent=2),
            fallback,
            quiet=True,
        )

        model_passed = str(review.get("passes_quality_check", "")).lower() == "true"
        if model_passed and not local_issues:
            return current

        current = {
            "linkedin_connect_note": limit_text(
                review.get("linkedin_connect_note") or safe_revision["linkedin_connect_note"], 279
            ),
            "cold_email_subject": str(
                review.get("cold_email_subject") or safe_revision["cold_email_subject"]
            ).strip(),
            "cold_email_body": str(
                review.get("cold_email_body") or safe_revision["cold_email_body"]
            ).strip(),
        }

    if outreach_quality_issues(current["linkedin_connect_note"], current["cold_email_body"]):
        current = make_safe_outreach_revision(context)
    current["linkedin_connect_note"] = limit_text(current["linkedin_connect_note"], 279)
    return current


def outreach_quality_issues(linkedin_note, email_body):
    issues = []
    note = str(linkedin_note or "").strip()
    body = str(email_body or "").strip()
    combined = f"{note} {body}".lower()

    if not note or len(note) >= 280:
        issues.append("LinkedIn note must be under 280 characters")
    if len(note) < 25 or len(body) < 80:
        issues.append("Message is too short to sound natural and explain the value")
    if combined.count("!") > 1 or any(
        phrase in combined
        for phrase in ["act now", "limited time", "guaranteed", "once-in-a-lifetime", "buy now", "don't miss out"]
    ):
        issues.append("Message sounds salesy or spammy")
    if not any(word in combined for word in ["short-form", "vertical drama", "production", "content", "partnership"]):
        issues.append("Value proposition is not clear")
    if body.count("?") != 1:
        issues.append("Email must have one simple call to action")
    if any(
        phrase in combined
        for phrase in ["proven results", "industry-leading", "award-winning", "we have helped", "i know you need"]
    ):
        issues.append("Message may contain an unsupported claim")
    return issues


def make_safe_outreach_revision(context):
    name = context.get("contact_name") or "there"
    company = context.get("company") or "your team"
    angle = context.get("suggested_outreach_angle") or "short-form vertical drama production"
    note = limit_text(
        f"Hi {name}, I'm exploring {angle} and thought {company} could be relevant. I'd be glad to connect and learn whether this is useful for your team.",
        279,
    )
    return {
        "linkedin_connect_note": note,
        "cold_email_subject": f"Short-form production idea for {company}",
        "cold_email_body": (
            f"Hi {name},\n\n"
            f"I came across {company} while researching short-form vertical drama and production partnerships. "
            f"I'm exploring {angle} and would value your perspective.\n\n"
            "Would you be open to a brief conversation next week?\n\n"
            "Best,\n"
            "Your Name"
        ),
    }


def limit_text(text, max_length):
    text = " ".join(str(text or "").split())
    if len(text) <= max_length:
        return text
    return text[: max_length - 3].rstrip() + "..."


def is_valid_email(email):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(email or "").strip()))


def is_ready_for_gmail(value):
    return str(value or "").strip().upper() == "TRUE"


def create_gmail_drafts(source_name):
    rows = read_csv_rows(source_name)
    approved_rows = []

    for row in rows:
        if is_sample_row(row) or not is_ready_for_gmail(row.get("ready_for_gmail")):
            continue
        email = row.get("recipient_email") or row.get("public_email")
        if is_valid_email(email):
            approved_rows.append((row, email))

    if not approved_rows:
        print("No Gmail drafts were created. Set ready_for_gmail to TRUE only on rows you have reviewed and approved.")
        return

    service = get_gmail_service()
    if service is None:
        print("Gmail is not configured, so no drafts were created. CSV files are still available for review.")
        return

    created_count = 0
    for row, email in approved_rows:
        subject = row.get("cold_email_subject") or f"Short-form vertical drama idea for {row.get('company', 'your team')}"
        body = row.get("cold_email_body") or make_basic_email(row)
        create_single_gmail_draft(service, email, subject, body)
        created_count += 1

    print(f"Created {created_count} Gmail drafts. Nothing was sent.")


def get_gmail_service():
    if Credentials is None or InstalledAppFlow is None or build is None:
        print("Gmail packages are not installed. Run: pip install -r requirements.txt")
        return None

    credentials_path = BASE_DIR / "credentials.json"
    token_path = BASE_DIR / "token.json"

    if not credentials_path.exists():
        print("credentials.json was not found. Follow the README Gmail setup steps first.")
        return None

    creds = None

    # Step 1: Reuse the local OAuth token if you have already approved Gmail draft access.
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), GMAIL_SCOPES)

    # Step 2: Refresh an expired token when Google allows it.
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

    # Step 3: If there is no valid token, open the Google OAuth approval screen.
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), GMAIL_SCOPES)
        creds = flow.run_local_server(port=0)

    # Step 4: Save the token locally so you do not need to approve every run.
    with token_path.open("w", encoding="utf-8") as token_file:
        token_file.write(creds.to_json())

    # Step 5: Build the Gmail API client. This code only creates drafts and never sends email.
    return build("gmail", "v1", credentials=creds)


def create_single_gmail_draft(service, to_email, subject, body):
    message = MIMEText(body)
    message["to"] = to_email
    message["subject"] = subject

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    draft_body = {"message": {"raw": raw_message}}
    service.users().drafts().create(userId="me", body=draft_body).execute()


def run_safe_test():
    """Check local setup only. This function never uses OpenAI or Gmail."""
    print("\nAI Opportunity Radar - Safe Setup Test")
    print("This test does not call OpenAI, create Gmail drafts, or send email.\n")
    needs_attention = False

    print("1. CSV files and headers")
    for filename, expected_headers in REQUIRED_CSV_HEADERS.items():
        path = BASE_DIR / filename
        if not path.exists():
            print(f"   [NEEDS ATTENTION] {filename} is missing.")
            needs_attention = True
            continue

        try:
            with path.open("r", newline="", encoding="utf-8-sig") as file:
                actual_headers = next(csv.reader(file), [])
        except OSError as error:
            print(f"   [NEEDS ATTENTION] {filename} could not be read: {error}")
            needs_attention = True
            continue

        if actual_headers == expected_headers:
            print(f"   [OK] {filename}")
        else:
            print(f"   [NEEDS ATTENTION] {filename} has different headers.")
            print(f"      Expected: {', '.join(expected_headers)}")
            print(f"      Found:    {', '.join(actual_headers) if actual_headers else '(no headers)'}")
            needs_attention = True

    print("\n2. Python requirements")
    requirements_path = BASE_DIR / "requirements.txt"
    if requirements_path.exists():
        print("   [OK] requirements.txt exists.")
    else:
        print("   [NEEDS ATTENTION] requirements.txt is missing.")
        needs_attention = True

    missing_packages = []
    for package_name, module_name in REQUIRED_PACKAGES.items():
        try:
            installed = importlib.util.find_spec(module_name) is not None
        except (ImportError, ModuleNotFoundError, ValueError):
            installed = False
        if not installed:
            missing_packages.append(package_name)

    if missing_packages:
        print(f"   [NEEDS ATTENTION] Missing: {', '.join(missing_packages)}")
        print("      Run: pip install -r requirements.txt")
        needs_attention = True
    else:
        print("   [OK] Required packages are installed.")

    print("\n3. Settings file")
    if (BASE_DIR / ".env").exists():
        print("   [OK] .env exists. The test did not read or use your API key.")
    else:
        print("   [NEEDS ATTENTION] .env is missing.")
        print("      Copy .env.example to .env, then add your OpenAI API key.")
        needs_attention = True

    print("\nResult")
    if needs_attention:
        print("Setup needs attention. Follow the messages above, then run: python main.py test")
    else:
        print("Setup looks ready. You can run: python main.py analyze or python main.py source")
    print("No OpenAI calls or Gmail actions were made.\n")
    return not needs_attention


def main():
    parser = argparse.ArgumentParser(description="AI Opportunity Radar")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("analyze", help="Analyze rows in opportunities.csv and save output.csv")
    subparsers.add_parser("source", help="Source up to 10 daily contacts and save daily_targets.csv")
    subparsers.add_parser(
        "linkedin-search",
        help="Create manual LinkedIn People search links without scraping LinkedIn",
    )
    subparsers.add_parser("test", help="Safely check files, headers, packages, and .env")

    draft_parser = subparsers.add_parser("draft-gmail", help="Create Gmail drafts from a reviewed CSV file")
    draft_parser.add_argument(
        "--from-file",
        default="daily_targets.csv",
        choices=["daily_targets.csv", "output.csv"],
        help="Choose which reviewed CSV file to use for Gmail drafts",
    )

    subparsers.add_parser("all", help="Run analyze and source")

    args = parser.parse_args()

    # Keep the safe test completely local: it does not load or use API credentials.
    if args.command != "test" and load_dotenv is not None:
        load_dotenv(BASE_DIR / ".env")

    if args.command == "analyze":
        analyze_opportunities()
    elif args.command == "source":
        source_daily_contacts()
    elif args.command == "linkedin-search":
        generate_linkedin_search_tasks()
    elif args.command == "test":
        run_safe_test()
    elif args.command == "draft-gmail":
        create_gmail_drafts(args.from_file)
    elif args.command == "all":
        analyze_opportunities()
        source_daily_contacts()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
