"""
HigherGov → HubSpot Importer (with Otter.ai)
=============================================
Fetches solicitations from HigherGov and imports them into HubSpot as Deals.
Cleans up outdated solicitations. Syncs Otter.ai meeting transcripts and
action items to matching HubSpot deals.

HOW OTTER MATCHING WORKS:
  Include the solicitation number anywhere in your Otter meeting title or notes.
  Example meeting titles:
    "Kickoff call W912DR-26-R-0001"
    "Agency Q&A - W912DR-26-R-0001 - April 27"
  The script extracts the solicitation number and links the meeting to the
  matching HubSpot deal automatically.

OTTER NOTE:
  Otter.ai has no official public API. This script uses the unofficial
  otterai-api library (github.com/gmchad/otterai-api) with your Otter
  username/password. As an alternative, Otter's Zapier integration can
  push transcripts via webhook — see ZAPIER ALTERNATIVE comment below.

Requirements:
    pip install requests hubspot-api-client otterai-api

Setup:
    1. HIGHERGOV_API_KEY     — from your HigherGov account
    2. HUBSPOT_ACCESS_TOKEN  — HubSpot > Settings > Integrations > Private Apps
       Required scopes: crm.objects.deals.read, crm.objects.deals.write
                        (No notes or tasks scopes needed)
    3. OTTER_EMAIL / OTTER_PASSWORD — your Otter.ai login credentials
    4. SAM_GOV_API_KEY       — optional, free at sam.gov/profile/details
"""

import requests
import time
import re
from datetime import datetime, timedelta
import hubspot
from hubspot.crm.deals import SimplePublicObjectInputForCreate, ApiException

# Otter unofficial API
try:
    from otterai import OtterAI
    OTTER_AVAILABLE = True
except ImportError:
    OTTER_AVAILABLE = False
    print("⚠️  otterai-api not installed. Run: pip install otterai-api")
    print("   Otter sync will be skipped.\n")


# ─── API Keys & Credentials ────────────────────────────────────────────────────

HIGHERGOV_API_KEY    = os.getenv("HIGHERGOV_API_KEY", "")
HUBSPOT_ACCESS_TOKEN = os.getenv("HUBSPOT_ACCESS_TOKEN", "")
SAM_GOV_API_KEY      = os.getenv("SAM_GOV_API_KEY", "")    # Optional

OTTER_EMAIL          = os.getenv("OTTER_EMAIL", "")
OTTER_PASSWORD       = os.getenv("OTTER_PASSWORD", "")


# ─── Opportunity Sources ───────────────────────────────────────────────────────

FETCH_FEDERAL     = True   # Federal opportunities via SAM.gov
FETCH_STATE_LOCAL = True   # State & local via HigherGov SLED endpoint
STATE_FILTER      = []     # e.g. ["TX", "VA"] — blank = all 50 states + DC

DAYS_BACK  = 7     # How many days back to pull new solicitations
PAGE_SIZE  = 100   # Max results per page (max 100)
NAICS_CODE = ""    # e.g. "541511" — federal only
SET_ASIDE  = ""    # e.g. "SBA", "8A" — federal only


# ─── HubSpot Settings ──────────────────────────────────────────────────────────

HUBSPOT_PIPELINE      = "default"
HUBSPOT_STAGE_NEW     = "appointmentscheduled"
HUBSPOT_STAGE_CLOSED  = "closedlost"
HUBSPOT_STAGE_AWARDED = "closedwon"
HUBSPOT_OWNER_ID      = ""

MAX_DESCRIPTION_CHARS = 65000
FETCH_FULL_DESC       = True

CLOSED_STATUSES  = {"cancelled", "canceled", "inactive", "closed", "withdrawn"}
AWARDED_STATUSES = {"awarded", "award", "complete", "completed"}


# ─── Otter Settings ────────────────────────────────────────────────────────────

# How many days back to look for Otter meetings to sync
OTTER_DAYS_BACK = 7

# Regex pattern to extract solicitation numbers from meeting titles/notes.
# Covers common federal formats: W912DR-26-R-0001, 36C10B25R0001, FA8501-25-R-001
# Adjust if your agency uses a different format.
OTTER_SOL_PATTERN = re.compile(
    r'\b([A-Z0-9]{2,10}[-][A-Z0-9]{2,6}[-][A-Z0-9]{1,2}[-][A-Z0-9]{2,6})\b'
    r'|\b([A-Z]{2,6}\d{4,}[A-Z]\d{4,})\b',
    re.IGNORECASE
)

# Max characters for transcript note in HubSpot (notes have a 65K char limit)
MAX_TRANSCRIPT_CHARS = 60000


# ─── Custom HubSpot Property Definitions ──────────────────────────────────────

CUSTOM_PROPERTIES = [
    {
        "name": "gov_market_level",
        "label": "Market Level",
        "description": "Federal or State/Local solicitation",
        "groupName": "dealinformation",
        "type": "enumeration",
        "fieldType": "select",
        "options": [
            {"label": "Federal",     "value": "Federal",     "displayOrder": 0, "hidden": False},
            {"label": "State/Local", "value": "State/Local", "displayOrder": 1, "hidden": False},
        ],
    },
    {
        "name": "gov_agency_name",
        "label": "Agency Name",
        "description": "The government agency that issued the solicitation",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_solicitation_number",
        "label": "Solicitation Number",
        "description": "Unique solicitation or RFP number",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_naics_code",
        "label": "NAICS Code",
        "description": "6-digit NAICS code (federal primarily)",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_set_aside_type",
        "label": "Set-Aside Type",
        "description": "Federal small business set-aside (e.g. SBA, 8(a), WOSB)",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_posted_date",
        "label": "Posted Date",
        "description": "Date the solicitation was originally posted",
        "groupName": "dealinformation",
        "type": "date", "fieldType": "date", "options": [],
    },
    {
        "name": "gov_source_url",
        "label": "Source URL",
        "description": "Direct link to the solicitation on SAM.gov or state portal",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_state",
        "label": "State",
        "description": "US state for the solicitation (State/Local only)",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_agency_type",
        "label": "Agency Type",
        "description": "Type of agency (e.g. Municipal, County, School District)",
        "groupName": "dealinformation",
        "type": "string", "fieldType": "text", "options": [],
    },
    {
        "name": "gov_status",
        "label": "Solicitation Status",
        "description": "Current status synced from HigherGov",
        "groupName": "dealinformation",
        "type": "enumeration",
        "fieldType": "select",
        "options": [
            {"label": "Active",    "value": "active",    "displayOrder": 0, "hidden": False},
            {"label": "Expired",   "value": "expired",   "displayOrder": 1, "hidden": False},
            {"label": "Cancelled", "value": "cancelled", "displayOrder": 2, "hidden": False},
            {"label": "Awarded",   "value": "awarded",   "displayOrder": 3, "hidden": False},
        ],
    },
    {
        "name": "gov_last_synced",
        "label": "Last Synced",
        "description": "Timestamp of the last HigherGov sync",
        "groupName": "dealinformation",
        "type": "date", "fieldType": "date", "options": [],
    },
    {
        "name": "gov_otter_last_synced",
        "label": "Otter Last Synced",
        "description": "Timestamp of the last Otter.ai meeting sync",
        "groupName": "dealinformation",
        "type": "date", "fieldType": "date", "options": [],
    },
    {
        "name": "gov_meeting_count",
        "label": "Meeting Count",
        "description": "Number of Otter.ai meetings linked to this deal",
        "groupName": "dealinformation",
        "type": "number", "fieldType": "number", "options": [],
    },
]


# ─── HigherGov Fetch Functions ─────────────────────────────────────────────────

def get_date_filter():
    return (datetime.today() - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")


def fetch_federal_page(page=1):
    endpoint = "https://www.highergov.com/api-external/opportunity/"
    params = {
        "api_key": HIGHERGOV_API_KEY, "captured_date": get_date_filter(),
        "page_size": PAGE_SIZE, "page_number": page, "source_type": "sam",
    }
    if NAICS_CODE: params["naics_code"]     = NAICS_CODE
    if SET_ASIDE:  params["set_aside_type"] = SET_ASIDE
    resp = requests.get(endpoint, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def fetch_state_local_page(page=1):
    endpoint = "https://www.highergov.com/api-external/opportunity/"
    params = {
        "api_key": HIGHERGOV_API_KEY, "captured_date": get_date_filter(),
        "page_size": PAGE_SIZE, "page_number": page,
        "source_type": "sled",
    }
    if STATE_FILTER:
        params["state"] = ",".join(STATE_FILTER)
    resp = requests.get(endpoint, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def fetch_all_from_endpoint(fetch_fn, label):
    all_results, page = [], 1
    while True:
        data    = fetch_fn(page=page)
        results = data.get("results", [])
        count   = data.get("count", 0)
        if not results:
            break
        for r in results:
            r["_market_level"] = label
        all_results.extend(results)
        print(f"    Page {page}: {len(results)} records (total: {len(all_results)}/{count})")
        if len(all_results) >= count:
            break
        page += 1
    return all_results


def fetch_all_opportunities():
    all_opps = []
    if FETCH_FEDERAL:
        print(f"📡 Fetching FEDERAL solicitations (last {DAYS_BACK} days)...")
        fed = fetch_all_from_endpoint(fetch_federal_page, "Federal")
        print(f"  → {len(fed)} federal records\n")
        all_opps.extend(fed)
    if FETCH_STATE_LOCAL:
        lbl = f"State/Local ({', '.join(STATE_FILTER)})" if STATE_FILTER else "State/Local (all states)"
        print(f"📡 Fetching {lbl} solicitations (last {DAYS_BACK} days)...")
        sled = fetch_all_from_endpoint(fetch_state_local_page, "State/Local")
        print(f"  → {len(sled)} state/local records\n")
        all_opps.extend(sled)
    print(f"✅ Total solicitations retrieved: {len(all_opps)}\n")
    return all_opps


def fetch_highergov_status(opportunity_id, is_sled=False):
    if not opportunity_id:
        return None
    try:
        base = "opportunity"
        resp = requests.get(
            f"https://www.highergov.com/api-external/{base}/{opportunity_id}/",
            params={"api_key": HIGHERGOV_API_KEY}, timeout=20,
        )
        if resp.status_code == 200:
            data = resp.json()
            for field in ("status", "opportunity_status", "active"):
                val = data.get(field)
                if val is not None:
                    return str(val).lower().strip()
    except Exception:
        pass
    return None


def fetch_highergov_description(opportunity_id, is_sled=False):
    if not opportunity_id:
        return None
    try:
        base = "opportunity"
        resp = requests.get(
            f"https://www.highergov.com/api-external/{base}/{opportunity_id}/",
            params={"api_key": HIGHERGOV_API_KEY}, timeout=20,
        )
        if resp.status_code == 200:
            data = resp.json()
            for field in ("description", "description_text", "opportunity_description", "summary", "body"):
                text = data.get(field, "")
                if text and len(text.strip()) > 20:
                    return clean_html(text.strip())
    except Exception:
        pass
    return None


def fetch_sam_description(notice_id, solicitation_number):
    if not SAM_GOV_API_KEY or SAM_GOV_API_KEY == "your-sam-gov-api-key-here":
        return None
    try:
        for param, val in [("noticeid", notice_id), ("solnum", solicitation_number)]:
            if not val:
                continue
            resp = requests.get(
                "https://api.sam.gov/opportunities/v2/search",
                params={"api_key": SAM_GOV_API_KEY, param: val,
                        "limit": 1, "includeSections": "description"},
                timeout=20,
            )
            if resp.status_code == 200:
                opps = resp.json().get("opportunitiesData", [])
                if opps:
                    desc = opps[0].get("description", "")
                    if desc and len(desc.strip()) > 20:
                        return clean_html(desc.strip())
        time.sleep(0.1)
    except Exception:
        pass
    return None


def clean_html(text):
    if not text:
        return text
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">") \
               .replace("&nbsp;", " ").replace("&#39;", "'").replace("&quot;", '"')
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def get_full_description(opp):
    if not FETCH_FULL_DESC:
        return None
    is_sled    = opp.get("_market_level") == "State/Local"
    opp_id     = opp.get("opportunity_id") or opp.get("id")
    notice_id  = opp.get("notice_id") or opp.get("sam_notice_id")
    sol_number = opp.get("solicitation_number", "")
    desc = fetch_highergov_description(opp_id, is_sled=is_sled)
    if desc:
        return desc
    if not is_sled:
        desc = fetch_sam_description(notice_id, sol_number)
        if desc:
            return desc
    for field in ("description", "description_text", "summary", "synopsis", "body"):
        text = opp.get(field, "")
        if text and len(str(text).strip()) > 20:
            return clean_html(str(text).strip())
    return None


# ─── Otter.ai Functions ────────────────────────────────────────────────────────

def extract_solicitation_number(text):
    """
    Extract a solicitation number from a string (meeting title or notes).
    Returns the first match found, or None.
    """
    if not text:
        return None
    match = OTTER_SOL_PATTERN.search(text)
    if match:
        return (match.group(1) or match.group(2)).upper().strip()
    return None


def fetch_otter_meetings():
    """
    Connect to Otter.ai and fetch recent meetings with transcripts/summaries.
    Returns a list of dicts with: title, transcript, summary, action_items,
    solicitation_number, date.
    """
    if not OTTER_AVAILABLE:
        return []

    print("🎙️  Connecting to Otter.ai...\n")
    try:
        otter = OtterAI()
        otter.login(OTTER_EMAIL, OTTER_PASSWORD)
    except Exception as e:
        print(f"  ❌ Otter login failed: {e}")
        return []

    cutoff = datetime.utcnow() - timedelta(days=OTTER_DAYS_BACK)
    meetings = []

    try:
        # Fetch speech list (all recordings)
        speeches = otter.get_speeches()
        speech_list = speeches.get("speeches", []) if isinstance(speeches, dict) else []

        print(f"  Found {len(speech_list)} total Otter recordings. Filtering last {OTTER_DAYS_BACK} days...\n")

        for speech in speech_list:
            # Filter by date
            created_ts = speech.get("created_at") or speech.get("start_time", 0)
            try:
                created_dt = datetime.utcfromtimestamp(int(created_ts))
            except Exception:
                continue
            if created_dt < cutoff:
                continue

            title = speech.get("title", "") or speech.get("meeting_name", "")

            # Try to get solicitation number from title first
            sol_number = extract_solicitation_number(title)

            # Fetch full speech detail for transcript + summary
            try:
                speech_id   = speech.get("otid") or speech.get("id")
                detail      = otter.get_speech(speech_id)
                detail_data = detail if isinstance(detail, dict) else {}

                # Extract transcript text
                transcript_parts = detail_data.get("speech_segments", [])
                transcript_text  = "\n".join(
                    f"[{seg.get('speaker_name', 'Speaker')}]: {seg.get('transcript', '')}"
                    for seg in transcript_parts
                    if seg.get("transcript")
                )

                # Extract summary
                summary = (
                    detail_data.get("summary")
                    or detail_data.get("outline")
                    or detail_data.get("abstract", "")
                )
                if isinstance(summary, list):
                    summary = "\n".join(str(s) for s in summary)

                # Extract action items
                action_items_raw = (
                    detail_data.get("action_items")
                    or detail_data.get("todos", [])
                )
                action_items = []
                for item in action_items_raw:
                    if isinstance(item, dict):
                        text = item.get("text") or item.get("transcript", "")
                        assignee = item.get("assignee", "")
                        action_items.append({"text": text, "assignee": assignee})
                    elif isinstance(item, str):
                        action_items.append({"text": item, "assignee": ""})

                # Try to get sol number from notes/description if not in title
                if not sol_number:
                    notes = detail_data.get("notes", "") or detail_data.get("description", "")
                    sol_number = extract_solicitation_number(notes)
                    # Last resort: scan the transcript itself
                    if not sol_number and transcript_text:
                        sol_number = extract_solicitation_number(transcript_text[:2000])

                if not sol_number:
                    print(f"  ⚠️  No solicitation number found in: '{title[:60]}' — skipping")
                    continue

                meetings.append({
                    "title":               title,
                    "date":                created_dt.strftime("%Y-%m-%d %H:%M UTC"),
                    "solicitation_number": sol_number,
                    "transcript":          transcript_text,
                    "summary":             str(summary) if summary else "",
                    "action_items":        action_items,
                    "otter_id":            speech_id,
                })
                print(f"  ✅ Found meeting: '{title[:50]}' → Sol# {sol_number}")

            except Exception as e:
                print(f"  ⚠️  Could not fetch detail for '{title[:50]}': {e}")

            time.sleep(0.3)  # Be polite to Otter's servers

    except Exception as e:
        print(f"  ❌ Error fetching Otter meetings: {e}")

    print(f"\n  → {len(meetings)} meetings matched to solicitation numbers.\n")
    return meetings


# ─── HubSpot Helper Functions ──────────────────────────────────────────────────

def init_hubspot_client():
    return hubspot.Client.create(access_token=HUBSPOT_ACCESS_TOKEN)


def parse_date_to_ms(date_str):
    """
    Convert YYYY-MM-DD to milliseconds epoch at midnight UTC.
    HubSpot date fields require exactly midnight UTC — any offset causes INVALID_DATE.
    """
    if not date_str:
        return None
    try:
        import calendar
        dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
        # Force midnight UTC using calendar.timegm (treats tuple as UTC, no local offset)
        midnight_utc_ms = calendar.timegm(dt.timetuple()) * 1000
        return str(midnight_utc_ms)
    except Exception:
        return None


def today_ms():
    """Today at midnight UTC in milliseconds."""
    import calendar
    today = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    return str(calendar.timegm(today.timetuple()) * 1000)


def setup_hubspot_properties(hs_client):
    print("🔧 Checking/creating custom HubSpot deal properties...\n")

    # Fetch existing properties
    try:
        resp     = hs_client.api_request({"path": "/crm/v3/properties/deals"})
        existing = {p["name"] for p in resp.json().get("results", [])}
    except Exception as e:
        print(f"  ⚠️  Could not fetch existing properties: {e}")
        existing = set()

    # Check if all custom properties already exist
    missing = [p for p in CUSTOM_PROPERTIES if p["name"] not in existing]
    if not missing:
        print(f"  ✓  All {len(CUSTOM_PROPERTIES)} custom properties already exist.\n")
        return

    # Try creating missing properties
    # Requires crm.schemas.deals.write — if 403, print manual instructions instead
    created_count = 0
    needs_manual  = []

    for prop in missing:
        try:
            resp = hs_client.api_request({
                "path": "/crm/v3/properties/deals", "method": "POST", "body": prop,
            })
            if hasattr(resp, "status_code") and resp.status_code == 403:
                needs_manual.append(prop)
            else:
                print(f"  ✅ Created: {prop['label']} ({prop['name']})")
                created_count += 1
        except Exception as e:
            if any(x in str(e).lower() for x in ["403", "forbidden", "scope"]):
                needs_manual.append(prop)
            else:
                print(f"  ❌ Failed to create {prop['name']}: {e}")
        time.sleep(0.2)

    if needs_manual:
        print(f"\n  ⚠️  Could not auto-create {len(needs_manual)} properties.")
        print("  Add crm.schemas.deals.write scope to your Private App, OR create manually:")
        print("  HubSpot → Settings → Data Management → Properties → Deals → Create property\n")
        for prop in needs_manual:
            print(f"     • {prop['label']} | name: {prop['name']} | type: {prop.get('fieldType', prop.get('type', 'text'))}")
        print("\n  The script will continue — custom fields will be blank until properties exist.\n")
    else:
        print(f"\n  → {created_count} new properties created.\n")


def map_opportunity_to_deal(opp, full_description=None):
    market_level = (opp.get("_market_level") or "Federal")
    is_sled      = market_level == "State/Local"
    sol_number   = (opp.get("solicitation_number") or "")
    title        = (opp.get("title") or "Unnamed Solicitation")
    market_tag   = "[FED]" if not is_sled else "[S&L]"
    deal_name    = f"{market_tag} [{sol_number}] {title}" if sol_number else f"{market_tag} {title}"

    properties = {
        "dealname":                deal_name[:255],
        "pipeline":                HUBSPOT_PIPELINE,
        "dealstage":               HUBSPOT_STAGE_NEW,
        "description":             (str(full_description) if full_description else "(Full description not available — check Source URL)")[:MAX_DESCRIPTION_CHARS],
        "gov_market_level":        safe_str(market_level, 50),
        "gov_agency_name":         safe_str(opp.get("agency_name"), 255),
        "gov_solicitation_number": safe_str(sol_number, 255),
        "gov_source_url":          safe_str(opp.get("url"), 255),
        "gov_status":              "active",
        "gov_last_synced":         today_ms(),
        "gov_meeting_count":       0,

    }

    deadline = opp.get("response_deadline") or opp.get("due_date")
    close_ms = parse_date_to_ms(deadline)
    if close_ms:
        properties["closedate"] = close_ms

    posted_ms = parse_date_to_ms(opp.get("posted_date"))
    if posted_ms:
        properties["gov_posted_date"] = posted_ms

    if is_sled:
        properties["gov_state"]       = safe_str(opp.get("state"), 100)
        properties["gov_agency_type"] = safe_str(opp.get("agency_type"), 255)
    else:
        properties["gov_naics_code"]     = safe_str(opp.get("naics_code"), 50)
        properties["gov_set_aside_type"] = safe_str(opp.get("set_aside_type"), 100)

    if HUBSPOT_OWNER_ID:
        properties["hubspot_owner_id"] = HUBSPOT_OWNER_ID

    return properties


def safe_str(val, max_len=None):
    """Convert any value to a safe string, handling None and non-string types."""
    result = str(val).strip() if val is not None else ""
    if max_len:
        result = result[:max_len]
    return result


def find_deal_by_solicitation_number(hs_client, sol_number):
    """
    Find a HubSpot deal by solicitation number.
    Returns the deal dict (id + properties) or None.
    """
    if not sol_number:
        return None
    try:
        resp = hs_client.api_request({
            "path": "/crm/v3/objects/deals/search",
            "method": "POST",
            "body": {
                "filterGroups": [{"filters": [
                    {"propertyName": "gov_solicitation_number",
                     "operator": "EQ", "value": sol_number},
                ]}],
                "properties": ["dealname", "gov_meeting_count", "gov_solicitation_number"],
                "limit": 1,
            },
        })
        results = resp.json().get("results", [])
        return results[0] if results else None
    except Exception:
        return None


def deal_already_exists(hs_client, sol_number, market_level):
    try:
        filters = [{"propertyName": "gov_market_level", "operator": "EQ", "value": market_level}]
        if sol_number:
            filters.append({"propertyName": "gov_solicitation_number", "operator": "EQ", "value": sol_number})
        resp = hs_client.api_request({
            "path": "/crm/v3/objects/deals/search",
            "method": "POST",
            "body": {"filterGroups": [{"filters": filters}], "limit": 1},
        })
        return resp.json().get("total", 0) > 0
    except Exception:
        return False


# Standard HubSpot deal fields — always safe to write without custom properties
STANDARD_DEAL_FIELDS = {
    "dealname", "pipeline", "dealstage", "description",
    "closedate", "amount", "hubspot_owner_id",
}

def create_hubspot_deal(hs_client, properties):
    """
    Create a deal. If it fails with 403 (likely due to missing custom gov_
    properties), retry using only standard HubSpot fields so the deal is
    still created — custom fields can be added later once properties exist.
    """
    try:
        input_data = SimplePublicObjectInputForCreate(properties=properties)
        return hs_client.crm.deals.basic_api.create(
            simple_public_object_input_for_create=input_data
        )
    except ApiException as e:
        if e.status == 400:
            print(f"       ❌ 400 Bad Request: {e.body}")
            raise
        if e.status == 403:
            # Retry with standard fields only
            safe_props = {k: v for k, v in properties.items() if k in STANDARD_DEAL_FIELDS}
            # Fold key gov_ fields into description as a fallback
            gov_summary = []
            for key in ("gov_market_level","gov_agency_name","gov_solicitation_number",
                        "gov_naics_code","gov_set_aside_type","gov_posted_date",
                        "gov_source_url","gov_state","gov_agency_type","gov_status"):
                val = properties.get(key, "")
                if val:
                    label = key.replace("gov_", "").replace("_", " ").title()
                    gov_summary.append(f"{label}: {val}")
            if gov_summary:
                header = "\n".join(gov_summary)
                existing_desc = safe_props.get("description", "")
                safe_props["description"] = (header + "\n\n" + existing_desc)[:65000]
            print(f"       ⚠️  Retrying without custom fields (add crm.schemas.deals.write scope to enable them)")
            input_data = SimplePublicObjectInputForCreate(properties=safe_props)
            return hs_client.crm.deals.basic_api.create(
                simple_public_object_input_for_create=input_data
            )
        raise


def update_hubspot_deal(hs_client, deal_id, properties):
    hs_client.api_request({
        "path": f"/crm/v3/objects/deals/{deal_id}",
        "method": "PATCH",
        "body": {"properties": properties},
    })


def append_otter_to_deal_description(hs_client, deal_id, current_desc, meeting):
    """
    Appends meeting transcript, summary, and action items directly to the
    deal description field. No notes or tasks scopes required.
    """
    title      = meeting["title"]
    date       = meeting["date"]
    sol_number = meeting["solicitation_number"]

    # Build the meeting block to append
    parts = [
        "",
        "=" * 60,
        f"📅 MEETING: {title}",
        f"🗓️  Date:    {date}",
        f"📋 Sol #:   {sol_number}",
        "=" * 60,
    ]

    if meeting.get("summary"):
        parts += ["", "── SUMMARY ──", meeting["summary"]]

    if meeting.get("action_items"):
        parts += ["", "── ACTION ITEMS ──"]
        for idx, item in enumerate(meeting["action_items"], 1):
            text     = item.get("text", "").strip()
            assignee = item.get("assignee", "")
            if text:
                line = f"  {idx}. {text}"
                if assignee:
                    line += f" (@{assignee})"
                parts.append(line)

    if meeting.get("transcript"):
        parts += ["", "── FULL TRANSCRIPT ──", meeting["transcript"]]

    meeting_block = "\n".join(parts)

    # Combine with existing description, respecting HubSpot's 65K char limit
    separator  = "\n\n" if current_desc else ""
    new_desc   = (current_desc or "") + separator + meeting_block
    if len(new_desc) > MAX_DESCRIPTION_CHARS:
        # Trim from the oldest content (top) to keep latest meeting at bottom
        new_desc = "...[older content trimmed]...\n" + new_desc[-(MAX_DESCRIPTION_CHARS - 30):]

    update_hubspot_deal(hs_client, deal_id, {"description": new_desc})
    return True


def flag_no_deadline_in_description(hs_client, deal_id, current_desc, deal_name):
    """
    Appends a ⚠️ flag to the deal description when no response deadline is set.
    No tasks scope required.
    """
    flag = (
        "\n\n" + "=" * 60 + "\n"
        "⚠️  ACTION REQUIRED: No response deadline set.\n"
        "Please review this solicitation on HigherGov or SAM.gov\n"
        "and update the Close Date field manually.\n"
        + "=" * 60
    )
    new_desc = ((current_desc or "") + flag)[:MAX_DESCRIPTION_CHARS]
    update_hubspot_deal(hs_client, deal_id, {"description": new_desc})


# ─── Otter → HubSpot Sync ──────────────────────────────────────────────────────

def sync_otter_to_hubspot(hs_client, meetings):
    """
    For each Otter meeting:
      1. Find the matching HubSpot deal by solicitation number
      2. Append summary, action items, and transcript to deal description
      3. Update gov_meeting_count and gov_otter_last_synced on the deal
    No notes or tasks scopes required — uses deal description only.
    """
    if not meetings:
        print("No Otter meetings to sync.\n")
        return

    synced   = 0
    no_match = 0
    errors   = 0

    print(f"🔗 Syncing {len(meetings)} Otter meetings to HubSpot...\n")

    for i, meeting in enumerate(meetings, 1):
        sol_number = meeting["solicitation_number"]
        title      = meeting["title"]

        deal = find_deal_by_solicitation_number(hs_client, sol_number)
        if not deal:
            print(f"  [{i}] ⚠️  No deal found for Sol# {sol_number} — '{title[:45]}'")
            no_match += 1
            continue

        deal_id    = deal["id"]
        deal_name  = deal.get("properties", {}).get("dealname", sol_number)
        meet_count = int(deal.get("properties", {}).get("gov_meeting_count") or 0)

        # Fetch current description so we can append to it
        try:
            resp = hs_client.api_request({
                "path": f"/crm/v3/objects/deals/{deal_id}",
                "method": "GET",
            })
            current_desc = resp.json().get("properties", {}).get("description", "") or ""
        except Exception:
            current_desc = ""

        print(f"  [{i}] 🔗 Matched: '{title[:40]}' → {deal_name[:40]}")

        try:
            # Append meeting content to deal description
            append_otter_to_deal_description(hs_client, deal_id, current_desc, meeting)
            action_count = len([a for a in meeting.get("action_items", []) if a.get("text", "").strip()])
            print(f"       ✅ Appended: summary + transcript + {action_count} action items")

            # Update deal metadata
            update_hubspot_deal(hs_client, deal_id, {
                "gov_meeting_count":     meet_count + 1,
                "gov_otter_last_synced": today_ms(),
            })

            synced += 1

        except Exception as e:
            print(f"  [{i}] ❌ Error syncing '{title[:50]}': {e}")
            errors += 1

        time.sleep(0.2)

    print(f"""
{'='*60}
  Otter Sync Complete
  ✅ Synced    : {synced}
  ⚠️  No match : {no_match}
  ❌ Errors    : {errors}
{'='*60}
""")


# ─── Cleanup Functions ─────────────────────────────────────────────────────────

def fetch_open_hubspot_deals(hs_client):
    print("🔍 Fetching open deals from HubSpot for cleanup...\n")
    all_deals, after = [], None
    while True:
        body = {
            "filterGroups": [{"filters": [
                {"propertyName": "pipeline",         "operator": "EQ",     "value": HUBSPOT_PIPELINE},
                {"propertyName": "dealstage",        "operator": "NOT_IN", "values": [HUBSPOT_STAGE_CLOSED, HUBSPOT_STAGE_AWARDED]},
                {"propertyName": "gov_market_level", "operator": "HAS_PROPERTY"},
            ]}],
            "properties": ["dealname", "dealstage", "closedate", "description",
                           "gov_solicitation_number", "gov_market_level", "gov_status"],
            "limit": 100,
        }
        if after:
            body["after"] = after
        try:
            resp    = hs_client.api_request({"path": "/crm/v3/objects/deals/search", "method": "POST", "body": body})
            data    = resp.json()
            results = data.get("results", [])
            all_deals.extend(results)
            after   = data.get("paging", {}).get("next", {}).get("after")
            if not after or not results:
                break
        except Exception as e:
            print(f"  ⚠️  Error fetching open deals: {e}")
            break
    print(f"  → Found {len(all_deals)} open deals to review.\n")
    return all_deals


def resolve_deal_status(deal):
    props      = deal.get("properties", {})
    sol_number = props.get("gov_solicitation_number", "")
    is_sled    = props.get("gov_market_level", "Federal") == "State/Local"
    close_date = props.get("closedate")

    if not close_date:
        return "no_deadline"

    try:
        close_dt = datetime.utcfromtimestamp(int(close_date) / 1000)
        if close_dt.date() < datetime.utcnow().date():
            hg_status = fetch_highergov_status(sol_number, is_sled=is_sled)
            if hg_status:
                if any(s in hg_status for s in AWARDED_STATUSES): return "awarded"
                if any(s in hg_status for s in CLOSED_STATUSES):  return "cancelled"
            return "expired"
    except Exception:
        pass

    hg_status = fetch_highergov_status(sol_number, is_sled=is_sled)
    if hg_status:
        if any(s in hg_status for s in AWARDED_STATUSES): return "awarded"
        if any(s in hg_status for s in CLOSED_STATUSES):  return "cancelled"

    return "ok"


def cleanup_hubspot_deals(hs_client):
    open_deals = fetch_open_hubspot_deals(hs_client)
    if not open_deals:
        print("No open deals to review.\n")
        return

    expired_count = cancelled_count = awarded_count = flagged_count = ok_count = 0
    print(f"🧹 Running cleanup on {len(open_deals)} open deals...\n")

    for i, deal in enumerate(open_deals, 1):
        props     = deal.get("properties", {})
        deal_id   = deal.get("id")
        deal_name = props.get("dealname", "Unknown Deal")
        action    = resolve_deal_status(deal)

        if action == "expired":
            update_hubspot_deal(hs_client, deal_id, {
                "dealstage": HUBSPOT_STAGE_CLOSED, "gov_status": "expired", "gov_last_synced": today_ms(),
            })
            print(f"  [{i}] 📅 Expired  → Closed Lost: {deal_name[:60]}")
            expired_count += 1

        elif action == "cancelled":
            update_hubspot_deal(hs_client, deal_id, {
                "dealstage": HUBSPOT_STAGE_CLOSED, "gov_status": "cancelled", "gov_last_synced": today_ms(),
            })
            print(f"  [{i}] ❌ Cancelled → Closed Lost: {deal_name[:60]}")
            cancelled_count += 1

        elif action == "awarded":
            update_hubspot_deal(hs_client, deal_id, {
                "dealstage": HUBSPOT_STAGE_AWARDED, "gov_status": "awarded", "gov_last_synced": today_ms(),
            })
            print(f"  [{i}] 🏆 Awarded  → Closed Won:  {deal_name[:60]}")
            awarded_count += 1

        elif action == "no_deadline":
            current_desc = props.get("description", "") or ""
            # Only flag once — check if already flagged
            if "ACTION REQUIRED: No response deadline" not in current_desc:
                flag_no_deadline_in_description(hs_client, deal_id, current_desc, deal_name)
                print(f"  [{i}] ⚠️  No deadline (flagged in description): {deal_name[:50]}")
            else:
                print(f"  [{i}] ⚠️  No deadline (already flagged): {deal_name[:55]}")
            update_hubspot_deal(hs_client, deal_id, {"gov_last_synced": today_ms()})
            flagged_count += 1

        else:
            print(f"  [{i}] ✓  Still active: {deal_name[:65]}")
            ok_count += 1

        time.sleep(0.2)

    print(f"""
{'='*60}
  Cleanup Complete
  📅 Expired  → Closed Lost : {expired_count}
  ❌ Cancelled → Closed Lost : {cancelled_count}
  🏆 Awarded  → Closed Won  : {awarded_count}
  ⚠️  No deadline (flagged) : {flagged_count}
  ✓  Still active           : {ok_count}
{'='*60}
""")


# ─── Import Function ───────────────────────────────────────────────────────────

def import_to_hubspot(opportunities, hs_client):
    created = skipped = errors = 0
    print(f"🚀 Importing {len(opportunities)} solicitations into HubSpot...\n")

    for i, opp in enumerate(opportunities, 1):
        sol_number   = (opp.get("solicitation_number") or "")
        title        = (opp.get("title") or "Unnamed Solicitation")
        market_level = (opp.get("_market_level") or "Federal")
        market_tag   = "[FED]" if market_level == "Federal" else "[S&L]"
        deal_name    = (f"{market_tag} [{sol_number}] {title}" if sol_number else f"{market_tag} {title}")[:255]

        try:
            if deal_already_exists(hs_client, sol_number, market_level):
                print(f"  [{i}] ⏭  Skipped (exists): {deal_name[:65]}")
                skipped += 1
                continue

            print(f"  [{i}] 🔍 Fetching description: {deal_name[:55]}...")
            full_desc  = get_full_description(opp)
            properties = map_opportunity_to_deal(opp, full_description=full_desc)
            create_hubspot_deal(hs_client, properties)

            # Flag no-deadline deals directly in description — no tasks scope needed
            if not (opp.get("response_deadline") or opp.get("due_date")):
                deal_resp = hs_client.api_request({
                    "path": "/crm/v3/objects/deals/search", "method": "POST",
                    "body": {"filterGroups": [{"filters": [
                        {"propertyName": "gov_solicitation_number", "operator": "EQ", "value": sol_number},
                        {"propertyName": "gov_market_level",        "operator": "EQ", "value": market_level},
                    ]}], "properties": ["description"], "limit": 1},
                })
                results = deal_resp.json().get("results", [])
                if results:
                    current_desc = results[0].get("properties", {}).get("description", "") or ""
                    flag_no_deadline_in_description(hs_client, results[0]["id"], current_desc, deal_name)
                print(f"  [{i}] ✅ Created + ⚠️ flagged (no deadline): {deal_name[:43]}")
            else:
                print(f"  [{i}] ✅ Created: {deal_name[:65]}")
            created += 1

        except ApiException as e:
            if e.status == 403:
                print(f"  [{i}] ❌ 403 Forbidden — check your Private App scopes include:")
                print(f"         crm.objects.deals.read + crm.objects.deals.write")
            else:
                print(f"  [{i}] ❌ HubSpot error: {e.status} - {e.reason}")
            errors += 1
        except Exception as e:
            print(f"  [{i}] ❌ Unexpected error: {e}")
            errors += 1

        time.sleep(0.15)

    print(f"""
{'='*60}
  Import Complete
  ✅ Created : {created}
  ⏭  Skipped : {skipped}
  ❌ Errors  : {errors}
{'='*60}
""")


# ─── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        hs_client = init_hubspot_client()

        # Step 1: Ensure all custom HubSpot properties exist
        setup_hubspot_properties(hs_client)

        # Step 2: Pull new solicitations from HigherGov and import as deals
        opportunities = fetch_all_opportunities()
        if opportunities:
            import_to_hubspot(opportunities, hs_client)
        else:
            print("No new solicitations found. Adjust your filters and try again.\n")

        # Step 3: Sync Otter.ai meeting transcripts and action items to deals
        if OTTER_AVAILABLE:
            meetings = fetch_otter_meetings()
            sync_otter_to_hubspot(hs_client, meetings)
        else:
            print("⏭️  Skipping Otter sync (library not installed).\n")

        # Step 4: Clean up outdated deals
        cleanup_hubspot_deals(hs_client)

    except requests.exceptions.HTTPError as e:
        print(f"❌ HTTP Error: {e.response.status_code} - {e.response.text}")
    except requests.exceptions.ConnectionError:
        print("❌ Connection error. Check your internet connection.")
    except requests.exceptions.Timeout:
        print("❌ Request timed out.")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
