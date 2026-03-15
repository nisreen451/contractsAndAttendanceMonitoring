import pandas as pd
from datetime import datetime, timedelta
import os

# =========================
# FILE PATHS
# =========================
attendance_file = r"C:\Users\CFWP\Desktop\Attendance Test\UpdatedCodes\InputFiles\Feb\AttendanceReport_252202691822.xlsx"
contracts_file = r"C:\Users\CFWP\Desktop\Attendance Test\UpdatedCodes\InputFiles\Feb\Contracts_252202691841.xlsx"
assigned_officers_file = r"C:\Users\CFWP\Desktop\Attendance Test\UpdatedCodes\InputFiles\Feb\AssignedAttendanceOfficersReport.xls (1).xls"

# =========================
# READ MONTH AND YEAR
# =========================
month_year = pd.read_excel(attendance_file, header=None, usecols="A,B", nrows=5)
month = month_year.iloc[4, 0]
year = month_year.iloc[4, 1]

if isinstance(month, str):
    month = datetime.strptime(month, "%B").month
else:
    month = int(month)
year = int(year)

# =========================
# DEFINE DATES
# =========================
today = datetime.today().date()
yesterday = today - timedelta(days=1)
month_start = datetime(year, month, 1).date()
month_end = datetime(year, month, pd.Period(f"{year}-{month}").days_in_month).date()

# =========================
# READ ATTENDANCE
# =========================
attendance_df = pd.read_excel(attendance_file, skiprows=5)
attendance_df.rename(columns={
    attendance_df.columns[0]: "Contract No.",
    attendance_df.columns[1]: "Contractor Name"
}, inplace=True)
attendance_df.columns = attendance_df.columns.str.strip()

all_cols = attendance_df.columns.tolist()
day_cols_old = all_cols[2:-4]
summary_cols_old = all_cols[-4:]

# Rename day columns into real dates
day_cols_new = []
for i, col in enumerate(day_cols_old, start=1):
    try:
        date_obj = datetime(year, month, i)
        new_name = date_obj.strftime("%d/%m/%Y %A")
    except ValueError:
        new_name = col
    day_cols_new.append(new_name)

attendance_df.rename(columns=dict(zip(day_cols_old, day_cols_new)), inplace=True)
attendance_df.rename(columns=dict(zip(summary_cols_old, ["UnConf.", "Abs.", "Conf.", "Hol."])), inplace=True)
attendance_df.drop(columns=["Contractor Name"], inplace=True)

# =========================
# READ CONTRACTS
# =========================
contracts_df = pd.read_excel(contracts_file)
contracts_df.columns = contracts_df.columns.str.strip()
contracts_df["Start Date"] = pd.to_datetime(contracts_df["Start Date"], errors="coerce").dt.date
contracts_df["Expiry Date"] = pd.to_datetime(contracts_df["Expiry Date"], errors="coerce").dt.date

# =========================
# ENSURE SAME TYPE FOR MERGE
# =========================
attendance_df["Contract No."] = attendance_df["Contract No."].astype(str)
contracts_df["Contract No."] = contracts_df["Contract No."].astype(str)

# =========================
# MERGE ATTENDANCE WITH CONTRACTS
# =========================
contracts_with_attendance = contracts_df.merge(
    attendance_df,
    on="Contract No.",
    how="left"
)

# =========================
# READ OFFICERS
# =========================
officers_df = pd.read_excel(assigned_officers_file, skiprows=13)
officers_df.columns = officers_df.columns.str.strip().str.replace('\n','').str.replace('\r','').str.replace('\xa0','')
officers_df = officers_df[['Email', 'Department', 'Area', 'Station']]

# =========================
# NORMALIZE TEXT
# =========================
for col in ["Department", "Area", "Station"]:
    contracts_with_attendance[col] = contracts_with_attendance[col].astype(str).str.strip().str.upper()
    officers_df[col] = officers_df[col].astype(str).str.strip().str.upper()

# =========================
# EXCLUDE EMAILS
# =========================
excluded_emails = {"s.badran@unrwa.org", "abd.abuamer@unrwa.org"}
officers_df["Email"] = officers_df["Email"].astype(str).str.strip().str.lower()
officers_df = officers_df[~officers_df["Email"].isin(excluded_emails)]

# =========================
# AGGREGATE EMAILS
# =========================
officers_agg = (
    officers_df.dropna(subset=["Email"])
    .groupby(["Department", "Area", "Station"], as_index=False)
    .agg({"Email": lambda x: ", ".join(sorted(set(x)))})
)

final_df = contracts_with_attendance.merge(
    officers_agg,
    on=["Department", "Area", "Station"],
    how="left"
)

# =========================
# ATTENDANCE LOGIC
# =========================
day_cols_dates = {
    col: datetime.strptime(col.split(" ")[0], "%d/%m/%Y").date()
    for col in final_df.columns
    if "/" in col and any(d in col for d in ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"])
}
sorted_day_cols = sorted(day_cols_dates.items(), key=lambda x: x[1])

def is_attended(v):
    return not pd.isna(v) and str(v).strip().upper() == "X"

def is_recorded(v):
    return not pd.isna(v) and str(v).strip().upper() in ("X", "A")

def is_project_beit_jala(row):
    return row["Department"] == "PROJECT" and row["Area"] == "BEIT JALA"

# =========================
# EMPTY ATTENDANCE
# =========================
def no_attendance_from_start_to_yesterday(row):
    start_date = row["Start Date"]
    expiry_date = row["Expiry Date"]

    # Ignore missing dates
    if pd.isna(start_date) or pd.isna(expiry_date):
        return False

    # EXCLUDE contracts where Start Date = Expiry Date
    if start_date == expiry_date:
        return False

    # Ignore future contracts
    if start_date > yesterday:
        return False

    # Check if there was any attendance between start date and yesterday
    for col, col_date in day_cols_dates.items():
        if start_date <= col_date <= yesterday and is_attended(row[col]):
            return False

    return True


final_df["No Attendance (Start → Yesterday)"] = final_df.apply(no_attendance_from_start_to_yesterday, axis=1)

# =========================
# 7 CONTINUOUS EMPTY DAYS
# =========================
def has_7_continuous_no_attendance(row):
    start_date = row["Start Date"]
    expiry_date = row["Expiry Date"]

    if pd.isna(start_date) or pd.isna(expiry_date):
        return False
    if (yesterday - start_date).days < 6:
        return False

    check_start = max(start_date, month_start)
    check_end = min(expiry_date, yesterday)

    if check_start > check_end:
        return False

    valid_days = [(col, d) for col, d in sorted_day_cols if check_start <= d <= check_end]

    for i in range(len(valid_days) - 6):
        window = valid_days[i:i+7]
        if (window[-1][1] - window[0][1]).days != 6:
            continue
        if all(not is_recorded(row[col]) for col, _ in window):
            return True
    return False

final_df["No Attendance 7 Continuous Days"] = final_df.apply(has_7_continuous_no_attendance, axis=1)

# =========================
# FILTER ACTIVE CONTRACTS FOR THESE SHEETS
# =========================
def is_active_in_month(row):
    start = row["Start Date"]
    end = row["Expiry Date"]
    if pd.isna(start) or pd.isna(end):
        return False
    if str(row["Status"]).strip().lower() == "canceled":
        return False
    # Check if contract period overlaps with this month
    return (start <= month_end) and (end >= month_start)

active_df = final_df[final_df.apply(is_active_in_month, axis=1)].copy()
empty_attendance_df = active_df[active_df["No Attendance (Start → Yesterday)"]].copy()
no_attendance_7cont_df = active_df[active_df["No Attendance 7 Continuous Days"]].copy()

# =========================
# OVER 5 DAYS & WEEKEND ATTENDANCE
# =========================

def over_5_days_per_week(row):
    streak = 0
    flagged_ranges = []
    current_range = []

    sorted_days = sorted(day_cols_dates.items(), key=lambda x: x[1])

    for i, (col, current_date) in enumerate(sorted_days):
        if is_attended(row[col]):
            if streak == 0:
                current_range = [current_date]
            else:
                prev_date = sorted_days[i-1][1]
                if (current_date - prev_date).days == 1:
                    current_range.append(current_date)
                else:
                    streak = 0
                    current_range = [current_date]

            streak += 1

            # More than 5 consecutive attendance days is an error
            if streak > 5:
                flagged_ranges.append(
                    f"{current_range[0].strftime('%d/%m')} → {current_range[-1].strftime('%d/%m')}"
                )
        else:
            streak = 0
            current_range = []

    return " | ".join(flagged_ranges)


# Any attendance on Saturday is an error
def get_saturday_attendance(row):
    return ", ".join(
        col.split(" ")[0]
        for col in day_cols_dates
        if "Saturday" in col and is_attended(row[col])
    )


# Any attendance on Friday is an error
def get_friday_attendance(row):
    return ", ".join(
        col.split(" ")[0]
        for col in day_cols_dates
        if "Friday" in col and is_attended(row[col])
    )


final_df["Over 5 Days Per Week"] = final_df.apply(over_5_days_per_week, axis=1)
final_df["Saturday Attendance"] = final_df.apply(get_saturday_attendance, axis=1)
final_df["Friday Attendance"] = final_df.apply(get_friday_attendance, axis=1)


# Attendance error if ANY of the rules are violated
def attendance_error_flag(row):
    return "YES" if (
        row["Over 5 Days Per Week"] or
        row["Saturday Attendance"] or
        row["Friday Attendance"]
    ) else "NO"


final_df["Attendance Error Flag"] = final_df.apply(attendance_error_flag, axis=1)


# =========================
# KPI COMPARISON
# =========================
contracts_df["Contract No."] = pd.to_numeric(contracts_df["Contract No."], errors="coerce")
final_df["Contract No."] = pd.to_numeric(final_df["Contract No."], errors="coerce")

kpi_df = pd.DataFrame({
    "Metric": [
        "Contracts Count (Original File)",
        "Contracts Count (Resulted File)",
        "Contracts ID Sum (Original File)",
        "Contracts ID Sum (Resulted File)"
    ],
    "Value": [
        contracts_df["Contract No."].nunique(),
        final_df["Contract No."].nunique(),
        contracts_df["Contract No."].sum(),
        final_df["Contract No."].sum()
    ]
})

# =========================
# SAVE OUTPUT IN SAME INPUT FOLDER
# =========================
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# Path to the 'OutputFiles' folder inside UpdatedCodes
output_folder = r"C:\Users\CFWP\Desktop\Attendance Test\UpdatedCodes\OutputFiles"

# Ensure the folder exists (optional, in case of typo)
if not os.path.exists(output_folder):
    raise FileNotFoundError(f"Output folder does not exist: {output_folder}")

# Output file path
output_file = os.path.join(
    output_folder,
    f"contractsWithAttendanceMonitoring_{month}_{year}_{timestamp}.xlsx"
)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Contracts Data", index=False)
    kpi_df.to_excel(writer, sheet_name="KPIs_Comparison", index=False)
    empty_attendance_df.to_excel(writer, sheet_name="Empty_Attendance", index=False)
    no_attendance_7cont_df.to_excel(writer, sheet_name="NoAttendance_7Days", index=False)

print(f"FINAL file created successfully in:\n{output_file}")
print(f"Empty Attendance Records: {len(empty_attendance_df)}")
print(f"7 Continuous No Attendance Records (Active Only): {len(no_attendance_7cont_df)}")