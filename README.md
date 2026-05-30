# Payroll Comparer Streamlit App

This app compares two payroll files by Employee UID and exports only mismatched rows.

## Run locally

1. Install Python 3.10+
2. In this folder, run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload these files:
   - `app.py`
   - `requirements.txt`
3. Go to Streamlit Community Cloud.
4. Create a new app from the GitHub repo.
5. Set the main file path to `app.py`.
6. Deploy.

Users will then open the Streamlit URL, upload two files, compare, and download `Mismatches.xlsx`.

## Notes

- Supported input files: `.csv`, `.xlsx`, `.xls`
- Apple Numbers files should be exported to Excel first.
- Mismatches are aligned by Employee UID.
- Duplicate UIDs are reported separately.
