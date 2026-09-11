# CNIC Attendance Separator - Online Deployment

This version adds:
- Login protection
- Environment-based username/password
- Secret Flask session key
- Production Gunicorn server
- Render deployment configuration
- Permanent CNIC PDF loaded from `/etc/secrets/cnic_data.pdf` on Render
- Attendance PDF upload and copy-number matching
- Excel and TXT downloads

## Important security note
Do NOT commit or push `cnic_data.pdf` to GitHub. It contains sensitive CNIC information.
On Render, upload it as a Secret File named `cnic_data.pdf`. The included `.gitignore` also ignores this file.

## Render settings
Create a Web Service from this project.

Build command:
`pip install -r requirements.txt`

Start command:
`gunicorn app:app`

Environment variables:
- `LOGIN_USERNAME` = your login username
- `LOGIN_PASSWORD` = a strong password
- `SECRET_KEY` = long random secret (or let Render generate it)
- `CNIC_PDF_PATH` = `/etc/secrets/cnic_data.pdf`
- `COOKIE_SECURE` = `1`

Secret File:
- Filename: `cnic_data.pdf`
- Contents: upload/paste your permanent CNIC PDF

## Local run
Put `cnic_data.pdf` beside `app.py`, then:

```bash
python -m venv venv
venv\\Scripts\\activate
pip install -r requirements.txt
set LOGIN_USERNAME=admin
set LOGIN_PASSWORD=your-password
set SECRET_KEY=your-long-random-secret
python app.py
```

Open `http://127.0.0.1:5000`.
