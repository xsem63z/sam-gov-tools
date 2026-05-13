# Changelog — generate_proposal.py

All notable changes to the proposal generator are documented here.

## [Unreleased]

## [1.0.0] — 2026-05-13

### Added
- Initial release of `generate_proposal.py`
- Fetches solicitation details from SAM.gov by notice ID using the v2 opportunities API
- Fallback search scan across up to 5 years of postings when the detail endpoint returns no data
- Follows description URLs returned by SAM.gov to retrieve full solicitation text
- Generates a 10-section FAR-compliant proposal via Claude (`claude-opus-4-5`)
  - Cover Page, Executive Summary, Understanding of Requirements, Technical Approach,
    Management Approach, Key Personnel, Past Performance, Price/Cost Narrative,
    Certifications & Representations, Conclusion
- Exports proposal as a styled `.docx` using Node.js and the `docx` library
  - Heading styles, bullet list formatting, page breaks, cover page
- Falls back to plain `.txt` output if Node.js or the `docx` package is unavailable
- CLI arguments: `--notice`, `--company`, `--uei`, `--cage`, `--capabilities`,
  `--past-perf`, `--personnel`, `--output`, `--desc`, `--title`
- `--desc` flag to bypass SAM.gov description fetch and supply text directly
- `--title` flag to override the opportunity title in the generated document
- Loads API keys from environment variables or a `.env` file via `python-dotenv`
- HTML tag stripping and whitespace normalization for all fetched description text
