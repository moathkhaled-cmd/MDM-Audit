Added scripts to generate audit report, call OpenAI with a strict prompt, produce CSV and Excel, and send the report by email.

Files added:
- scripts/audit_report.py  # main auditing script
- scripts/email_report.py  # send report via SMTP
- prompts/audit_prompt.txt # strict AI prompt to return only 0/1
- requirements.txt

Usage:
1. Prepare input CSV at data/products.csv. It must have at least these columns in order (or containing these names):
   - product_name
   - brochure_path  (local path or URL; empty if none)
   - official_site  (URL of official website; optional)
   - feature columns (one column per feature, e.g. "MDM Enroll", "Remote Wipe", ...)

2. (Optional) Create a .env file with OPENAI_API_KEY and SMTP credentials.

3. Run:
   pip install -r requirements.txt
   python3 scripts/audit_report.py --input data/products.csv --out data/report.xlsx
   python3 scripts/email_report.py --file data/report.xlsx

Notes:
- The AI prompt forces the model to reply with only comma-separated 0/1 values so downstream processing stays deterministic.
- Rows audited from official website but without brochure are highlighted light yellow in the Excel and have a note.
