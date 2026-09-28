# Attendance Monitoring Report Generator

Faculty web app to convert **ERP Attendance Monitoring exports** into the standard departmental format for **First / Second / Third** attendance monitoring.

## Features

- Upload ERP attendance Excel (wide format with D / A / % per subject)
- Choose monitoring period: **FIRST**, **SECOND**, or **THIRD**
- Set minimum attendance criteria (default 75%)
- Auto-detect class info, subjects, and faculty
- Preview critical students (any course below criteria)
- Download formatted report matching the departmental template:
  - Subject-wise % only
  - Overall %
  - Count of courses below criteria
  - **CRITICAL** flag
  - Summary statistics
  - Observations / signature blocks

## Quick start (local)

```bash
# 1. Clone this repo
git clone https://github.com/<your-username>/attendance-monitoring-app.git
cd attendance-monitoring-app

# 2. Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

Open **http://localhost:8501** in your browser.

## Deploy on Streamlit Community Cloud (free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with GitHub → **New app**
4. Select this repository, branch `main`, main file `app.py`
5. Click **Deploy**

Your faculty can then open the public URL and upload reports without installing anything.

## How to use

1. **Upload** the ERP Attendance Monitoring `.xlsx` file  
2. In the sidebar, select **FIRST / SECOND / THIRD** monitoring  
3. Optionally set minimum criteria (%) and Program Coordinator name  
4. Review detected class info and critical students preview  
5. Click **Generate Formatted Report** → **Download**

## Expected input format

ERP export should contain:

| Block | Content |
|-------|---------|
| Header | Branch, Department, Class Name, Division, Date range |
| Columns | Sr, Roll No, Student Name, then each subject as **D / A / %** (3 columns) |
| Rows | One row per student with delivered, attended, and percentage |
| Total | Overall attendance % column |

## Output format

Matches the standard **Sample of Attendance Report** structure used by the Department of CST:

- Title: `DEPARTMENT OF CST` + monitoring label  
- Subject columns with codes, type (CORE / ELECTIVE), faculty names  
- Overall %, count below criteria, Whether Critical  
- Summary: enrolled / below 75% / 70% / 65% / 60%  
- Observations, Action Taken, signature lines (Mentor, HoD, Dean, etc.)

## Project structure

```
attendance-monitoring-app/
├── app.py                 # Streamlit application
├── requirements.txt       # Python dependencies
├── README.md              # This file
└── .gitignore
```

## Requirements

- Python 3.9+
- streamlit
- openpyxl

## License

For internal academic use (Manav Rachna / MRU – Department of CST).
