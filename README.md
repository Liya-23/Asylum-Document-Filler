# Asylum Document Filler

Cleaned and hardened Tornado app for generating asylum newcomer, renewal, and confirmation documents from Word templates.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Place these Word templates next to `app.py`:

- `ASYLUM_NEW_COMER_Template.docx`
- `ASYLUM_RENEWAL_Template.docx`
- `CONFIRMATION_LETTER_Template.docx`

Each template can use these placeholders:

- `[Todays Date]`
- `[title]`
- `[full names]`
- `[DOB]`
- `[pronoun]`
- `[gender]`

## Run

```powershell
python app.py --port 2026
```

Open `http://localhost:2026`.

PDF generation uses LibreOffice if it is installed. If LibreOffice is missing, the DOCX is still generated and the app reports that PDF conversion is unavailable.
