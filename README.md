# Contracts with Attendance Monitoring

This project processes **contracts and attendance data** to generate a combined Excel report with KPIs and attendance alerts.

## Features

- Merge contracts with attendance records.
- Flag empty attendance records.
- Detect 7 continuous days of no attendance.
- Detect attendance errors: over 5 days/week or weekend attendance.
- KPI comparison between original and processed data.
- Outputs multiple sheets: `Contracts Data`, `KPIs_Comparison`, `Empty_Attendance`, `NoAttendance_7Days`.

## How to Run

1. Place input Excel files in `InputFiles/`.
2. Install dependencies:

```bash
pip install -r requirements.txt
