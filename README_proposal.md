# SAM.gov Government Contract Proposal Generator

Fetches a solicitation from SAM.gov by notice ID, generates a complete FAR-compliant proposal using Claude AI, and saves it as a formatted Word document.

## Requirements

**Python packages:**
```
pip install requests anthropic python-dotenv
```

**Node.js** (for `.docx` output — falls back to `.txt` if not available):
```
npm install docx
```

## Setup

1. Copy `.env.example` to `.env` (or export variables in your shell):
   ```
   SAM_API_KEY=your_sam_gov_api_key
   ANTHROPIC_API_KEY=your_anthropic_api_key
   ```

2. **SAM.gov API key** — free, register at [sam.gov/profile/details](https://sam.gov/profile/details)

3. **Anthropic API key** — available at [console.anthropic.com](https://console.anthropic.com)

## Usage

```bash
python3 generate_proposal.py --notice NOTICE_ID [options]
```

### Required argument

| Argument | Description |
|----------|-------------|
| `--notice` | SAM.gov notice ID (found in the opportunity URL) |

### Optional arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `--company` | `[Your Company Name]` | Your company name |
| `--uei` | `[UEI]` | Unique Entity Identifier |
| `--cage` | `[CAGE Code]` | CAGE code |
| `--capabilities` | `[To be provided]` | Core capabilities narrative |
| `--past-perf` | `[To be provided]` | Past performance summary |
| `--personnel` | `[To be provided]` | Key personnel description |
| `--output` | Auto-named by solicitation # | Output file path |
| `--desc` | *(fetched from SAM.gov)* | Paste description text directly (skips SAM.gov fetch) |
| `--title` | *(from SAM.gov)* | Override the opportunity title |

### Examples

```bash
# Minimal — placeholders used for company info
python3 generate_proposal.py --notice 27c272eff8bb4025be4ff71fbabdd890

# Full company info
python3 generate_proposal.py \
  --notice 27c272eff8bb4025be4ff71fbabdd890 \
  --company "Acme Realty LLC" \
  --uei ABC123XYZ \
  --cage 1A2B3 \
  --capabilities "Property management, facility operations, GSA lease administration" \
  --past-perf "USDA Denver lease management (2021-2024), EPA Region 8 facilities (2020-2023)" \
  --personnel "Jane Smith, PMP — Project Manager; John Doe — Facilities Lead"

# Provide description directly (useful if SAM.gov API is slow or you have the text handy)
python3 generate_proposal.py \
  --notice 27c272eff8bb4025be4ff71fbabdd890 \
  --desc "The agency requires lease administration services for..."
```

## Output

The script generates a `.docx` file named `proposal_<SOLICITATION_NUMBER>_<DATE>.docx` in the current directory (or the path specified by `--output`). If Node.js/`docx` is unavailable, a `.txt` fallback is saved instead.

The proposal includes 10 sections:

1. Cover Page
2. Executive Summary
3. Understanding of Requirements
4. Technical Approach
5. Management Approach
6. Key Personnel
7. Past Performance
8. Price/Cost Narrative
9. Certifications & Representations
10. Conclusion

## How it works

1. **Fetch** — Calls the SAM.gov detail endpoint. If that returns no data, falls back to a paginated search scan across up to 5 years of postings.
2. **Extract description** — Pulls the description field; if it's a URL, fetches the content from that URL.
3. **Generate** — Sends solicitation metadata + description + your company info to Claude (`claude-opus-4-5`) and receives a structured proposal.
4. **Export** — Parses the proposal into sections and renders a styled Word document via a Node.js script using the `docx` library.

## After running

1. Open the Word document and review all sections.
2. Fill in any `[placeholder]` fields with real data.
3. Have a subject matter expert review the technical approach.
4. Submit before the response deadline listed on SAM.gov.

> **Note:** AI-generated proposals are a strong starting point but should always be reviewed, tailored, and verified by a qualified proposal writer before submission.
