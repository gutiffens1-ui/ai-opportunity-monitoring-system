# AI Opportunity Radar

AI Opportunity Radar is a beginner-friendly Python tool for finding, evaluating, and organizing short-form vertical drama production and business development opportunities.

It can:

- analyze opportunities you add to `opportunities.csv`
- source up to 10 public or user-provided contacts from `seed_companies.csv`
- create manual LinkedIn People search tasks without opening or scraping LinkedIn
- improve LinkedIn notes and cold email drafts with a two-attempt quality loop
- save unique contacts in `contact_database.csv`
- optionally create reviewed Gmail drafts

It never sends email and never scrapes LinkedIn.

## First Thing To Do

On Windows:

1. Install Python from [python.org](https://www.python.org/downloads/). Select **Add Python to PATH** during installation.
2. Open PowerShell in this project folder.
3. Create a virtual environment:

```powershell
python -m venv .venv
```

4. Turn it on:

```powershell
.\.venv\Scripts\Activate.ps1
```

5. Install the required packages:

```powershell
pip install -r requirements.txt
```

6. Copy `.env.example` to `.env`:

```powershell
Copy-Item .env.example .env
```

7. Open `.env` and replace `your_openai_api_key_here` with your OpenAI API key.
8. Run the safe setup test:

```powershell
python main.py test
```

The test only checks local files, CSV headers, installed packages, and `.env`. It never calls OpenAI, creates Gmail drafts, or sends email.

## CSV Files To Edit First

Edit these input files:

- `opportunities.csv`: companies or contacts you already want to evaluate.
- `seed_companies.csv`: companies, public pages, RSS feeds, and optional verified contacts to use for sourcing.

Rows containing `SAMPLE`, `EXAMPLE`, or `Replace With` are demonstrations. The script skips them. Replace those words when a row contains real information you want the script to process.

These files are generated or updated by the script:

- `output.csv`: final opportunity analysis and outreach drafts.
- `daily_targets.csv`: up to 10 contacts from the latest sourcing run.
- `contact_database.csv`: unique contacts collected across sourcing runs.
- `linkedin_search_tasks.csv`: manual LinkedIn People search links generated from seed companies.

Generated outreach rows always start with `ready_for_gmail` set to `FALSE`.

## Add Your OpenAI API Key

Your `.env` file should contain:

```text
OPENAI_API_KEY=your_real_api_key_here
OPENAI_MODEL=gpt-4o-mini
MIN_CONTACT_SCORE=7
```

`MIN_CONTACT_SCORE` controls the daily contact quality threshold. If it is missing or invalid, the script uses `7`.

Keep `.env` private. It is already included in `.gitignore`.

## Test Your Setup

Run this whenever you want to check the project safely:

```powershell
python main.py test
```

The report uses `[OK]` and `[NEEDS ATTENTION]` messages and tells you what to fix. It does not use your API key.

## Run Opportunity Analysis

1. Open `opportunities.csv` in Excel.
2. Replace one or more example rows with real information.
3. Keep the header names unchanged.
4. Run:

```powershell
python main.py analyze
```

The final results are saved in `output.csv`. Sample rows are skipped.

## Run Daily Contact Sourcing

In this project, **daily** means you run this command when you want a new daily list:

```powershell
python main.py source
```

The project does not run automatically unless you set up Windows Task Scheduler.

The script reads `seed_companies.csv`, saves up to 10 contacts scoring at least `MIN_CONTACT_SCORE`, updates `daily_targets.csv`, and appends unique contacts to `contact_database.csv`. If fewer than 10 good contacts are supported by the available evidence, it saves only those contacts and explains this in PowerShell.

If no contacts are found, the program recommends:

```powershell
py main.py linkedin-search
```

This creates manual LinkedIn search tasks so you can review possible contacts yourself and add verified contacts back into `seed_companies.csv`.

The included seed companies show useful categories:

- ReelShort and DramaBox for vertical drama
- My Drama for mobile-first short series
- Runway for AI video workflows and partnerships
- Pocket FM for serialized story IP and adaptations

### Add Seed Companies

Each `seed_companies.csv` row contains:

- `company`
- `company_website`
- `description`
- `target_contact_titles`, separated by semicolons
- `suggested_contacts`, for public contacts you already verified
- `public_search_links`, separated by semicolons
- `rss_feed_url`, if available
- `notes`

Use this format inside `suggested_contacts`:

```text
Name | Title | Public company or profile page | Public email
```

Do not add LinkedIn pages for scraping. You may keep a LinkedIn profile URL as a contact reference, but this script will not fetch it.

## Create Manual LinkedIn Search Tasks

Run:

```powershell
py main.py linkedin-search
```

This command reads `seed_companies.csv` and creates `linkedin_search_tasks.csv`. For every company, it creates searches for producer, production, content, partnerships, business development, founder, and CEO roles.

The command only builds ordinary LinkedIn People search URLs. It does not:

- open or scrape LinkedIn
- log in to LinkedIn
- collect private information
- send connection requests
- contact anyone

### Review The Search Tasks

1. Open `linkedin_search_tasks.csv` in Excel.
2. Find the `linkedin_search_url` column.
3. Click or press `Ctrl` while clicking a URL to open it in your normal browser. Sign in manually if LinkedIn asks.
4. Review the search results yourself. Confirm the person's name, title, and company using public information.
5. Update `manual_review_status`, for example `Good contact`, `Not a fit`, or `Needs more research`.
6. Add your own comments in `notes`.

When you find a good public contact, the preferred next step is to copy them into the matching company's `suggested_contacts` cell in `seed_companies.csv`:

```text
Name | Title | Public profile or company page | Public email if available
```

You may instead add a complete row to `daily_targets.csv`. Keep `ready_for_gmail` set to `FALSE` until you have reviewed and approved the email draft.

Running `linkedin-search` again refreshes the links while preserving existing review statuses and notes for the same company and title.

## Optional Public Website Reading And Search

Website reading is off by default. To allow public company pages you selected, update `.env`:

```text
FETCH_WEBSITES=true
FETCH_PUBLIC_LINKS=true
```

Only use public pages that permit access. Do not add private or login-only pages.

Optional live search uses SerpAPI. Add a key only if you choose to use that separate service:

```text
SERPAPI_API_KEY=your_serpapi_key_here
```

The project works without SerpAPI.

## Outreach Quality Loop

After opportunity analysis or contact sourcing, the script checks each LinkedIn note and email draft. It checks note length, natural tone, sales pressure, clear value, one email call to action, unsupported claims, and spammy language.

If a draft fails, it is revised. The loop allows no more than two revision attempts, and only the final version is saved.

## Review CSV Files In Excel

Open `output.csv` and `daily_targets.csv` before using any outreach text. Open `linkedin_search_tasks.csv` when doing manual LinkedIn research.

For easier reading in Excel:

1. Select the whole sheet by clicking the small triangle above row `1` and left of column `A`.
2. Open **Home > Format > AutoFit Column Width**.
3. Use **Home > Format > AutoFit Row Height** if an email body is clipped.
4. Turn on **Wrap Text** for the email and explanation columns if helpful.

The script writes generated CSV files in an Excel-friendly UTF-8 format.

## Avoid Duplicate Contacts

Before adding a sourced contact to `contact_database.csv`, the script compares:

- `contact_name`
- `company`
- `contact_link`

The same combination is not added twice. Demonstration rows are also removed when the contact database is updated.

## Optional Windows Task Scheduler

You can schedule `source` once per day. This is optional.

1. First confirm that `python main.py source` works when you run it yourself.
2. Open **Task Scheduler** from the Windows Start menu.
3. Click **Create Basic Task**.
4. Name it `AI Opportunity Radar Daily Source`.
5. Choose **Daily** and select a time.
6. Choose **Start a program**.
7. For **Program/script**, enter:

```text
C:\Users\16564\OneDrive\Desktop\ai-opportunity-radar\.venv\Scripts\python.exe
```

8. For **Add arguments**, enter:

```text
main.py source
```

9. For **Start in**, enter:

```text
C:\Users\16564\OneDrive\Desktop\ai-opportunity-radar
```

10. Finish the task and use **Run** once to test it.

The scheduled task only performs sourcing and CSV updates. It does not create Gmail drafts or send emails. Your computer must be on and connected to the internet when the task runs.

## Gmail Drafts Setup

Gmail support is optional. It uses OAuth and the minimum draft-creation permission:

```text
https://www.googleapis.com/auth/gmail.compose
```

A Gmail draft is created only when both conditions are true:

- `recipient_email` or `public_email` is valid
- you manually change `ready_for_gmail` to `TRUE`

Any other value, including blank or `FALSE`, is not approved. The script never sends email.

### Enable The Gmail API

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create or select a project.
3. Search for **Gmail API** and click **Enable**.
4. Open **APIs & Services > OAuth consent screen**.
5. Configure the consent screen and add your Gmail address as a test user if Google asks.
6. Open **APIs & Services > Credentials**.
7. Click **Create Credentials > OAuth client ID**.
8. Choose **Desktop app**.
9. Download the JSON file and rename it `credentials.json`.
10. Place it in this project folder beside `main.py`.

The first draft run opens Google's OAuth approval page and stores the local token in `token.json`. Both `credentials.json` and `token.json` are ignored by Git.

### Create Reviewed Gmail Drafts

1. Review `output.csv` or `daily_targets.csv` in Excel.
2. Change `ready_for_gmail` to `TRUE` only for rows you approve.
3. Save the CSV.
4. Run one command:

```powershell
python main.py draft-gmail --from-file daily_targets.csv
```

or:

```powershell
python main.py draft-gmail --from-file output.csv
```

Open Gmail and select **Drafts** to inspect every message. You must send any message yourself from Gmail.

## Project Files

- `main.py`: commands, OpenAI analysis, quality checks, sourcing, CSV handling, and Gmail draft safety.
- `opportunities.csv`: opportunity input examples.
- `seed_companies.csv`: sourcing input examples.
- `output.csv`: analyzed opportunities.
- `daily_targets.csv`: latest sourced contacts.
- `contact_database.csv`: unique historical contacts.
- `linkedin_search_tasks.csv`: clickable LinkedIn People searches for manual review.
- `requirements.txt`: Python packages.
- `.env.example`: safe settings template.
- `.gitignore`: excludes private credentials, tokens, and local environment files.

## Safety Summary

- No automatic email sending.
- No Gmail drafts without `ready_for_gmail=TRUE` and a valid email.
- No LinkedIn scraping.
- No automatic LinkedIn login, browsing, profile collection, or connection requests.
- No processing of rows marked `SAMPLE`, `EXAMPLE`, or `Replace With`.
- Human review comes before every Gmail draft.
