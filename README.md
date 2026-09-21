# 📊 Sistem Penyelarasan Kehadiran Bulanan (HR Attendance Reconciler)

Streamlit app untuk menyelaraskan rekod kehadiran staf, laporan cuti, dan time-off jabatan ke dalam satu matriks Excel.

## Features

- Upload PDF clock-in/out attendance per department
- Upload company leave report (all staff)
- Upload time-off/timeslip reports per department
- Generates a reconciled Excel matrix with colour-coded attendance status

## Usage

1. Select year and month
2. Upload attendance PDFs (one per department)
3. Upload leave report PDF
4. Upload time-off PDFs (optional)
5. Click **Jana Laporan Final Excel** to generate and download the report

## Running Locally

```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

Hosted on [Streamlit Community Cloud](https://streamlit.io/cloud).

## Requirements

- Python 3.10+
- See `requirements.txt`
