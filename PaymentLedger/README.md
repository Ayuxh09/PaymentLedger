# PaymentLedger

HTML/CSS/Python payment ledger with OCR and Excel database.

## Views
- `/` — User view: upload screenshot + enter purchase.
- `/owner/login` — Owner login.
- `/owner` — Owner dashboard and payment totals.
- `/owner/download` — Download Excel database.

## Default development credentials
Username: `owner`
Password: `change-this-password`

Change these with environment variables before real use:
`OWNER_USERNAME`, `OWNER_PASSWORD`, `FLASK_SECRET_KEY`.

## Run
```bash
pip install -r requirements.txt
python app.py
```
Tesseract OCR must also be installed on the machine.

Excel schema:
Record ID, Date, Time, Status, Paid To, UPI ID, Amount, Transaction ID, Debited From, UTR, What Did You Purchase?, Created At.
