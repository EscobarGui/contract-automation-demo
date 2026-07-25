# Contract Automation — Demo

Small demo of the type of automation I built in production at a fintech company: turning a manual, template-based contract creation process into a script that validates incoming data and generates ready-to-send Word documents automatically.

**Real-world result this is modeled on:** reduced contract processing time from ~20 minutes to ~1 minute 20 seconds per contract (~93% reduction) by automating client data entry, document generation, and logging.

## What it does

1. Reads incoming contract requests from a CSV (`contract_requests.csv`) — this simulates a form submission, CRM export, or intake spreadsheet.
2. Validates each request (missing fields, invalid dates, invalid values) and skips/reports anything incomplete instead of silently generating a bad contract.
3. Generates a formatted `.docx` service agreement for every valid request, with parties, commercial terms, and standard clauses filled in automatically.
4. Logs every request (generated or skipped, with reason) to `processing_log.csv`.
5. Prints a summary comparing automated processing time against the manual baseline.

## Run it

```bash
node generate-contracts.js
```

Generated contracts land in `generated_contracts/`, and the run log is written to `processing_log.csv`.

## Why this matters for a client

This is the exact shape of problem a lot of small businesses have: someone on the team spends real hours per week manually filling in the same document template with slightly different data each time, with no validation and no record of what was processed. This script replaces that with something that runs in seconds, catches bad data before it becomes a bad contract, and leaves an audit trail.

## Adapting this for a real client

- Swap the CSV source for a Google Sheet, Airtable, form submission webhook, or CRM export.
- Swap the contract template/clauses for the client's actual agreement.
- Add e-signature integration (e.g. DocuSign API) as the next step after generation.
- Add email delivery of the generated contract directly to the client's inbox.
