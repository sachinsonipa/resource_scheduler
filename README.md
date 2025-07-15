
# Resource Scheduler WebApp

This Flask-based app allows you to:

- Manage Resources with daily work hour limits (Green, Yellow, Red)
- Add Work Items with estimated hours and a preferred resource
- Automatically compute expected end date and task status (On Track, At Risk, Very Risky)
- Store and retrieve data using Excel (`ResourceSheet.xlsx`)

## Installation

```bash
pip install flask pandas openpyxl
python app.py
```

## Usage

1. Open your browser at `http://127.0.0.1:5000`
2. Add resources and assign work items.
3. End date and status will be auto-calculated based on resource availability and holidays.
