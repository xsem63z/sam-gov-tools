"""
SAM.gov Government Contract Proposal Generator
-----------------------------------------------
Pulls a solicitation from SAM.gov by notice ID, generates a full
government contract proposal using Claude AI, and saves it as a
professionally formatted Word document (.docx).

SETUP:
1. pip install requests anthropic --break-system-packages
2. Fill in your API keys below
3. Run: python3 generate_proposal.py

USAGE:
  python3 generate_proposal.py --notice NOTICE_ID [options]

EXAMPLES:
  python3 generate_proposal.py --notice 27c272eff8bb4025be4ff71fbabdd890
  python3 generate_proposal.py --notice 27c272eff8bb4025be4ff71fbabdd890 --company "Acme Realty LLC" --uei ABC123
"""

import os
import re
import sys
import json
import argparse
import requests
import subprocess
import tempfile
import textwrap
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

# ─── Configuration ────────────────────────────────────────────────────────────

SAM_API_KEY       = os.getenv("SAM_API_KEY", "")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

SAM_DETAIL_URL    = "https://api.sam.gov/opportunities/v2/opportunities/{notice_id}"
ANTHROPIC_URL     = "https://api.anthropic.com/v1/messages"

# ─── SAM.gov Helpers ──────────────────────────────────────────────────────────

def clean_text(raw):
    """Strips HTML tags and normalizes whitespace."""
    if not raw:
        return ""
    clean = re.sub(r"<[^>]+>", " ", raw)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean


def fetch_description_from_url(url):
    """Fetches description content from a URL returned by SAM.gov."""
    try:
        params = {"api_key": SAM_API_KEY} if "api.sam.gov" in url else {}
        resp = requests.get(url, params=params, timeout=30)
        if resp.ok:
            return clean_text(resp.text)
    except requests.RequestException as e:
        print(f"  Warning: Could not fetch description URL: {e}")
    return ""


SAM_SEARCH_URL = "https://api.sam.gov/opportunities/v2/search"

def fetch_opportunity(notice_id):
    """Fetches full opportunity details from SAM.gov.
    Tries the detail endpoint first, then falls back to search endpoint."""
    print(f"Fetching SAM.gov opportunity: {notice_id}")

    # Try detail endpoint first
    url = SAM_DETAIL_URL.format(notice_id=notice_id)
    try:
        resp = requests.get(url, params={"api_key": SAM_API_KEY}, timeout=30)
        if resp.ok:
            data = resp.json()
            opps = data.get("opportunitiesData", [])
            if opps:
                return opps[0]
        print(f"  Detail endpoint unavailable ({resp.status_code}) — trying search endpoint...")
    except requests.RequestException as e:
        print(f"  Detail endpoint error: {e} — trying search endpoint...")

    # Fall back: SAM.gov search does not support noticeid filter directly.
    # Instead search recent windows and match notice ID from results.
    from datetime import timedelta
    today = datetime.today()
    found = None

    for years_back in range(0, 5):
        date_to   = today - timedelta(days=364 * years_back)
        date_from = date_to - timedelta(days=364)  # max 364 days — SAM.gov rejects exactly 365
        offset = 0
        print(f"  Scanning {date_from.strftime('%m/%d/%Y')} to {date_to.strftime('%m/%d/%Y')}...")
        while True:
            try:
                params = {
                    "api_key":    SAM_API_KEY,
                    "limit":      200,
                    "offset":     offset,
                    "postedFrom": date_from.strftime("%m/%d/%Y"),
                    "postedTo":   date_to.strftime("%m/%d/%Y"),
                }
                resp = requests.get(SAM_SEARCH_URL, params=params, timeout=30)
                resp.raise_for_status()
                data = resp.json()
                opps = data.get("opportunitiesData", [])
                total = data.get("totalRecords", 0)

                # Search for matching notice ID in this page
                for opp in opps:
                    if opp.get("noticeId", "").lower() == notice_id.lower():
                        print(f"  Found via search scan.")
                        return opp

                # Paginate if more records exist
                offset += len(opps)
                if offset >= total or not opps:
                    break

            except requests.RequestException as e:
                print(f"  Search window error: {e}")
                break

    print("  Error: Notice ID not found.")
    print(f"  Please verify at: https://sam.gov/opp/{notice_id}/view")
    print("  If the notice is valid, try providing the description text directly using --desc flag.")
    sys.exit(1)


def extract_description(opp):
    """Extracts the description text, following URLs if needed."""
    raw = opp.get("description", "")
    if not raw:
        return ""
    if raw.strip().lower().startswith("http"):
        print("  Description is a URL — fetching content...")
        return fetch_description_from_url(raw.strip())
    return clean_text(raw)


# ─── Proposal Generation ──────────────────────────────────────────────────────

def generate_proposal(opp, description, company_info):
    """Calls Claude to generate the proposal text."""
    print("Generating proposal with Claude AI...")

    title      = opp.get("title", "")
    sol_num    = opp.get("solicitationNumber", "")
    department = opp.get("fullParentPathName", "").split(".")[-1].strip()
    naics      = opp.get("naicsCode", "")
    set_aside  = opp.get("typeOfSetAsideDescription", "") or "None"
    deadline   = opp.get("responseDeadLine", "") or "See solicitation"
    posted     = opp.get("postedDate", "")

    prompt = f"""You are an expert government contract proposal writer with deep knowledge of FAR compliance, 
federal procurement, and winning proposal strategies.

Generate a complete, professional government contract proposal responding to this solicitation.
The proposal must be thorough, specific to the requirements, FAR-compliant, and written in a 
formal, confident tone.

SOLICITATION DETAILS:
- Title: {title}
- Solicitation Number: {sol_num}
- Agency: {department}
- NAICS Code: {naics}
- Set-Aside: {set_aside}
- Response Deadline: {deadline}
- Posted Date: {posted}

SOLICITATION DESCRIPTION:
{description or "See attached solicitation documents."}

OFFEROR INFORMATION:
- Company: {company_info['company']}
- UEI: {company_info['uei']}
- CAGE Code: {company_info['cage']}
- Capabilities: {company_info['capabilities']}
- Past Performance: {company_info['past_performance']}
- Key Personnel: {company_info['key_personnel']}

Generate a complete proposal with ALL of the following sections. Use clear section headers 
exactly as shown. Be specific and tie every section directly to the solicitation requirements.

---
SECTION 1: COVER PAGE
Include: company name, UEI, CAGE, solicitation number, date, and a one-paragraph executive summary.

SECTION 2: EXECUTIVE SUMMARY
2-3 paragraphs summarizing the company's qualifications and value proposition for this specific opportunity.

SECTION 3: UNDERSTANDING OF REQUIREMENTS
Demonstrate thorough understanding of the agency's needs as stated in the solicitation. Reference specific requirements.

SECTION 4: TECHNICAL APPROACH
Detailed methodology for how the company will fulfill each requirement. Include specific processes, tools, and deliverables.

SECTION 5: MANAGEMENT APPROACH
Organizational structure, project management methodology, quality control, and risk mitigation approach.

SECTION 6: KEY PERSONNEL
Qualifications and roles of key staff assigned to this contract.

SECTION 7: PAST PERFORMANCE
Relevant past contracts with similar scope, including outcomes and lessons learned.

SECTION 8: PRICE/COST NARRATIVE
General pricing approach, basis of estimate, and commitment to cost control. Do not include specific dollar amounts.

SECTION 9: CERTIFICATIONS & REPRESENTATIONS
Standard certifications including FAR 52.204-8 (Annual Representations), small business status if applicable, and compliance statements.

SECTION 10: CONCLUSION
Strong closing paragraph reaffirming commitment and fit for this opportunity.
---

Write the full proposal now. Be specific, professional, and compelling."""

    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    payload = {
        "model": "claude-opus-4-5-20251101",
        "max_tokens": 4000,
        "messages": [{"role": "user", "content": prompt}]
    }

    try:
        resp = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        text = "".join(b.get("text", "") for b in data.get("content", []))
        return text
    except requests.RequestException as e:
        print(f"  Error calling Claude API: {e}")
        sys.exit(1)


# ─── Word Document Generation ─────────────────────────────────────────────────

def save_as_docx(proposal_text, opp, output_path):
    """Saves the proposal as a formatted Word document using docx-js."""
    print("Creating Word document...")

    title   = opp.get("title", "Contract Proposal")
    sol_num = opp.get("solicitationNumber", "")
    date    = datetime.today().strftime("%B %d, %Y")

    # Parse sections from proposal text
    sections = []
    current_heading = None
    current_body = []

    for line in proposal_text.split("\n"):
        stripped = line.strip()
        if re.match(r"^SECTION \d+:", stripped) or re.match(r"^#{1,2} ", stripped):
            if current_heading:
                sections.append((current_heading, "\n".join(current_body).strip()))
            heading = re.sub(r"^#{1,2} ", "", stripped)
            heading = re.sub(r"^SECTION \d+:\s*", "", heading)
            current_heading = heading
            current_body = []
        else:
            current_body.append(line)

    if current_heading:
        sections.append((current_heading, "\n".join(current_body).strip()))

    # Build docx-js script
    js_sections = []

    # Cover page
    js_sections.append(f"""
        new Paragraph({{
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            children: [new TextRun({{ text: {json.dumps(title)}, bold: true }})]
        }}),
        new Paragraph({{
            alignment: AlignmentType.CENTER,
            spacing: {{ before: 240 }},
            children: [new TextRun({{ text: "Solicitation Number: {sol_num}", size: 24 }})]
        }}),
        new Paragraph({{
            alignment: AlignmentType.CENTER,
            spacing: {{ before: 120 }},
            children: [new TextRun({{ text: "Date: {date}", size: 24 }})]
        }}),
        new Paragraph({{
            children: [new PageBreak()]
        }}),
    """)

    # Proposal sections
    for heading, body in sections:
        js_sections.append(f"""
        new Paragraph({{
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({{ text: {json.dumps(heading)}, bold: true }})]
        }}),""")

        for para in body.split("\n\n"):
            para = para.strip()
            if not para:
                continue
            lines = para.split("\n")
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Detect bullet points
                if re.match(r"^[-•*]\s+", line):
                    line_text = re.sub(r"^[-•*]\s+", "", line)
                    js_sections.append(f"""
        new Paragraph({{
            numbering: {{ reference: "bullets", level: 0 }},
            children: [new TextRun({json.dumps(line_text)})]
        }}),""")
                else:
                    js_sections.append(f"""
        new Paragraph({{
            spacing: {{ before: 120, after: 120 }},
            children: [new TextRun({json.dumps(line)})]
        }}),""")

    js_content = "\n".join(js_sections)

    js_script = f"""
const {{ Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
         LevelFormat, PageBreak, WidthType }} = require('docx');
const fs = require('fs');

const doc = new Document({{
  numbering: {{
    config: [
      {{ reference: "bullets",
         levels: [{{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
           style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }} }}] }}
    ]
  }},
  styles: {{
    default: {{ document: {{ run: {{ font: "Arial", size: 24 }} }} }},
    paragraphStyles: [
      {{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
         run: {{ size: 28, bold: true, font: "Arial", color: "1F3864" }},
         paragraph: {{ spacing: {{ before: 360, after: 180 }}, outlineLevel: 0,
           border: {{ bottom: {{ style: "single", size: 6, color: "1F3864", space: 4 }} }} }} }},
      {{ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
         run: {{ size: 24, bold: true, font: "Arial", color: "2E75B6" }},
         paragraph: {{ spacing: {{ before: 240, after: 120 }}, outlineLevel: 1 }} }},
    ]
  }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 12240, height: 15840 }},
        margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }}
      }}
    }},
    children: [
      {js_content}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync({json.dumps(output_path)}, buffer);
  console.log("Document saved:", {json.dumps(output_path)});
}}).catch(err => {{
  console.error("Error:", err);
  process.exit(1);
}});
"""

    # Write and run JS script
    with tempfile.NamedTemporaryFile(suffix=".js", mode="w", delete=False) as f:
        f.write(js_script)
        js_path = f.name

    try:
        result = subprocess.run(
            ["node", js_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            print(f"  Word doc error: {result.stderr}")
            # Fall back to saving as text
            txt_path = output_path.replace(".docx", ".txt")
            with open(txt_path, "w") as f:
                f.write(proposal_text)
            print(f"  Saved as plain text instead: {txt_path}")
            return txt_path
    finally:
        os.unlink(js_path)

    return output_path


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate a government contract proposal from SAM.gov")
    parser.add_argument("--notice",       required=True,  help="SAM.gov notice ID")
    parser.add_argument("--company",      default="[Your Company Name]", help="Company name")
    parser.add_argument("--uei",          default="[UEI]",               help="UEI number")
    parser.add_argument("--cage",         default="[CAGE Code]",         help="CAGE code")
    parser.add_argument("--capabilities", default="[To be provided]",    help="Core capabilities")
    parser.add_argument("--past-perf",    default="[To be provided]",    help="Past performance")
    parser.add_argument("--personnel",    default="[To be provided]",    help="Key personnel")
    parser.add_argument("--output",       default=None,                  help="Output file path")
    parser.add_argument("--desc",         default=None,                  help="Paste description directly (skips SAM.gov description fetch)")
    parser.add_argument("--title",        default=None,                  help="Override opportunity title")
    args = parser.parse_args()

    print("\n" + "=" * 60)
    print("SAM.gov Government Contract Proposal Generator")
    print("=" * 60 + "\n")

    # Validate API keys
    if SAM_API_KEY == "YOUR_SAM_API_KEY_HERE":
        print("Error: Please set your SAM_API_KEY in the script or as an environment variable.")
        sys.exit(1)
    if ANTHROPIC_API_KEY == "YOUR_ANTHROPIC_API_KEY_HERE":
        print("Error: Please set your ANTHROPIC_API_KEY in the script or as an environment variable.")
        sys.exit(1)

    # Fetch opportunity
    opp = fetch_opportunity(args.notice)
    print(f"  Found: {opp.get('title', 'Unknown')}")
    print(f"  Agency: {opp.get('fullParentPathName', '').split('.')[-1].strip()}")
    print(f"  Solicitation: {opp.get('solicitationNumber', 'N/A')}")

    # Extract description
    description = args.desc if args.desc else extract_description(opp)
    if args.title:
        opp["title"] = args.title
    if description:
        print(f"  Description: {len(description)} characters retrieved")
    else:
        print("  Warning: No description found — proposal will be based on title/metadata only")

    # Company info
    company_info = {
        "company":          args.company,
        "uei":              args.uei,
        "cage":             args.cage,
        "capabilities":     args.capabilities,
        "past_performance": args.past_perf,
        "key_personnel":    args.personnel,
    }

    # Generate proposal
    proposal_text = generate_proposal(opp, description, company_info)
    print(f"  Proposal generated: {len(proposal_text)} characters")

    # Save output
    sol_num = opp.get("solicitationNumber", args.notice).replace("/", "-").replace(" ", "_")
    output_path = args.output or f"proposal_{sol_num}_{datetime.today().strftime('%Y%m%d')}.docx"
    final_path = save_as_docx(proposal_text, opp, output_path)

    print("\n" + "=" * 60)
    print(f"Proposal saved: {final_path}")
    print("=" * 60 + "\n")
    print("Next steps:")
    print("  1. Open the Word document and review all sections")
    print("  2. Fill in any [placeholder] fields with real data")
    print("  3. Have a subject matter expert review the technical approach")
    print("  4. Submit before the response deadline\n")


if __name__ == "__main__":
    main()
