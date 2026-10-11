# Filings Agent

Turns SEC filings into an analyst-style summary and a small three-scenario Excel model.

## What it does
1. **Fetches data from SEC EDGAR** (free): three years of XBRL financials, the latest 10-K/10-Q (MD&A and risk factors), and the latest earnings-release 8-K.
2. **Writes a summary** with an LLM: takeaways, results vs. prior period, drivers, guidance and tone, risks, and open questions.
3. **Builds an Excel model** (`openpyxl`) with bear/base/bull assumptions, a scenario switch, and live formulas for revenue, operating income, net income, EPS and free cash flow.

See [`examples/`](examples) for sample output (Apple, fiscal Q3 2026).

## Setup
```bash
pip install -r requirements.txt
```
Environment variables:
- `SEC_USER_AGENT` - your name and email (required by SEC), e.g. `"Jane Doe jane@example.com"`
- `ANTHROPIC_API_KEY` - only for `--mode api`

## Usage
```bash
python filings_agent.py AAPL                       # fetch + summarize via Claude API + build model
python filings_agent.py AAPL --transcript call.txt # include an earnings-call transcript
python filings_agent.py AAPL --mode dump           # fetch only; writes out/AAPL_context.json
python filings_agent.py AAPL --mode build          # build from out/AAPL_summary_body.md + out/AAPL_assumptions.json
```
`dump` / `build` let you write the summary and assumptions yourself, or with any LLM chat tool, without an API key. Outputs go to `out/`.

## Limitations
- US-GAAP filers only; unusual XBRL tags can leave gaps.
- Transcripts are not on EDGAR; supply your own file.
- The model is deliberately simple: net income = operating income x (1 - tax), no interest or other items.
- Scenario assumptions are model- or analyst-judgement, not company guidance. Not investment advice.
- The `--mode api` path has not been exercised end to end; the `dump` / `build` path has.

## Stack
Python, `requests`, `beautifulsoup4`, `openpyxl`, Anthropic API (optional).
