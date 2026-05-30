
import io
import re
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Payroll Comparer", page_icon="📊", layout="wide")


# -----------------------------
# Helpers
# -----------------------------

def normalize_col_name(value) -> str:
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_uid(value) -> Optional[str]:
    """Normalize Employee IDs so 1001136, 1001136.0, and ' 1001136 ' match."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text == "" or text.lower() in {"nan", "none", "null"}:
        return None
    text = text.replace(",", "")
    # Excel often reads whole-number IDs as floats.
    try:
        as_float = float(text)
        if as_float.is_integer():
            return str(int(as_float))
    except Exception:
        pass
    # Remove trailing .0 if present.
    text = re.sub(r"\.0$", "", text)
    return text


def to_number(series_or_value):
    return pd.to_numeric(series_or_value, errors="coerce").fillna(0)


def round_money(value):
    try:
        return round(float(value), 2)
    except Exception:
        return 0.0


def round_hours(value):
    try:
        return round(float(value), 2)
    except Exception:
        return 0.0


def find_header_row(raw_df: pd.DataFrame, required_terms: List[str], max_scan_rows: int = 30) -> int:
    """Find the row that most likely contains headers."""
    required_terms = [term.lower() for term in required_terms]
    best_row = 0
    best_score = -1

    for i in range(min(max_scan_rows, len(raw_df))):
        row_values = [normalize_col_name(v).lower() for v in raw_df.iloc[i].tolist() if not pd.isna(v)]
        row_joined = " | ".join(row_values)
        score = sum(1 for term in required_terms if term in row_joined)
        # Bonus for rows with more header-looking non-empty values.
        score += min(len(row_values), 20) / 100
        if score > best_score:
            best_score = score
            best_row = i

    return best_row


def read_excel_or_csv(uploaded_file, required_terms: List[str]) -> Tuple[pd.DataFrame, int]:
    """Read an uploaded Excel/CSV file and auto-detect the header row."""
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file, header=None, dtype=object)
        header_row = find_header_row(raw, required_terms)
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, header=header_row, dtype=object)
    else:
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)
        header_row = find_header_row(raw, required_terms)
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row, dtype=object)

    # Drop fully empty rows/cols and clean headers.
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [normalize_col_name(c) for c in df.columns]

    # Remove repeated header rows that sometimes appear in exported reports.
    if len(df) > 0:
        first_col = df.columns[0]
        df = df[df[first_col].astype(str).str.strip() != str(first_col).strip()]

    return df, header_row + 1


def has_columns(df: pd.DataFrame, cols: List[str]) -> bool:
    return all(c in df.columns for c in cols)


def detect_file_type(df: pd.DataFrame) -> str:
    """Detect whether this is the Workday earning register or Toast payroll export."""
    cols = {c.lower() for c in df.columns}

    if {"earning", "current period hours", "current period result"}.issubset(cols):
        return "workday"

    if {"regular hours", "overtime hours", "regular pay", "employee id"}.issubset(cols):
        return "toast"

    return "unknown"


def standardize_workday(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    required = ["Worker", "Employee ID", "Earning", "Current Period Hours", "Rate", "Current Period Result"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Workday file is missing required columns: {', '.join(missing)}")

    data = df.copy()
    data["uid"] = data["Employee ID"].apply(normalize_uid)
    data["employee_name"] = data["Worker"].astype(str).str.strip()
    data["earning_clean"] = data["Earning"].astype(str).str.strip()
    data["hours"] = pd.to_numeric(data["Current Period Hours"], errors="coerce").fillna(0)
    data["rate"] = pd.to_numeric(data["Rate"], errors="coerce")
    data["result"] = pd.to_numeric(data["Current Period Result"], errors="coerce").fillna(0)

    invalid = data[data["uid"].isna()][["employee_name", "Employee ID", "Earning", "Current Period Hours", "Rate", "Current Period Result"]].copy()

    valid = data[~data["uid"].isna()].copy()

    def agg_employee(group: pd.DataFrame) -> pd.Series:
        earning = group["earning_clean"].str.lower()

        regular = group[earning.eq("regular hourly pay")]
        overtime = group[earning.str.contains("overtime", na=False)]
        tips = group[earning.eq("tips charged")]

        # Use the regular hourly rate when available; otherwise use first non-null rate.
        regular_rates = regular["rate"].dropna()
        all_rates = group["rate"].dropna()

        return pd.Series({
            "Employee Name": group["employee_name"].dropna().astype(str).iloc[0] if len(group) else "",
            "Regular Hours": round_hours(regular["hours"].sum()),
            "Overtime Hours": round_hours(overtime["hours"].sum()),
            "Hourly Rate": round_money(regular_rates.iloc[0] if len(regular_rates) else (all_rates.iloc[0] if len(all_rates) else 0)),
            "Regular Pay": round_money(regular["result"].sum()),
            "Overtime Pay": round_money(overtime["result"].sum()),
            "Tips": round_money(tips["result"].sum()),
            "Total Pay": round_money(regular["result"].sum() + overtime["result"].sum()),
        })

    summary = valid.groupby("uid", dropna=False).apply(agg_employee).reset_index().rename(columns={"uid": "Employee ID"})
    return summary, invalid


def standardize_toast(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    required = ["Employee", "Employee ID", "Regular Hours", "Overtime Hours", "Hourly Rate", "Regular Pay", "Overtime Pay", "Total Pay"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Payroll export file is missing required columns: {', '.join(missing)}")

    data = df.copy()
    data["uid"] = data["Employee ID"].apply(normalize_uid)
    data["employee_name"] = data["Employee"].astype(str).str.strip()

    invalid = data[data["uid"].isna()][["employee_name", "Employee ID", "Regular Hours", "Overtime Hours", "Hourly Rate", "Regular Pay", "Overtime Pay", "Total Pay"]].copy()

    valid = data[~data["uid"].isna()].copy()

    # Prefer Non-Cash Tips because it matches Workday's "Tips Charged" in the sample file.
    tips_col = "Non-Cash Tips" if "Non-Cash Tips" in valid.columns else ("Total Tips" if "Total Tips" in valid.columns else None)

    grouped = valid.groupby("uid", dropna=False).agg({
        "employee_name": "first",
        "Regular Hours": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum(),
        "Overtime Hours": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum(),
        "Hourly Rate": lambda x: pd.to_numeric(x, errors="coerce").dropna().iloc[0] if len(pd.to_numeric(x, errors="coerce").dropna()) else 0,
        "Regular Pay": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum(),
        "Overtime Pay": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum(),
        "Total Pay": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum(),
    }).reset_index()

    if tips_col:
        tips = valid.groupby("uid")[tips_col].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).reset_index(name="Tips")
        grouped = grouped.merge(tips, on="uid", how="left")
    else:
        grouped["Tips"] = 0

    grouped = grouped.rename(columns={"uid": "Employee ID", "employee_name": "Employee Name"})
    for col in ["Regular Hours", "Overtime Hours"]:
        grouped[col] = grouped[col].apply(round_hours)
    for col in ["Hourly Rate", "Regular Pay", "Overtime Pay", "Total Pay", "Tips"]:
        grouped[col] = grouped[col].apply(round_money)

    return grouped, invalid


def compare_summaries(workday: pd.DataFrame, toast: pd.DataFrame) -> pd.DataFrame:
    compare_fields = [
        "Regular Hours",
        "Overtime Hours",
        "Hourly Rate",
        "Regular Pay",
        "Overtime Pay",
        "Total Pay",
        "Tips",
    ]

    merged = workday.merge(
        toast,
        on="Employee ID",
        how="outer",
        suffixes=(" - Workday", " - Payroll Export"),
        indicator=True,
    )

    rows = []

    for _, row in merged.iterrows():
        uid = row.get("Employee ID")
        workday_name = row.get("Employee Name - Workday", "")
        toast_name = row.get("Employee Name - Payroll Export", "")
        display_name = workday_name if pd.notna(workday_name) and str(workday_name).strip() else toast_name

        if row["_merge"] == "left_only":
            rows.append({
                "Employee ID": uid,
                "Employee Name": display_name,
                "Mismatch Type": "Missing employee",
                "Field": "Employee ID",
                "Workday Value": uid,
                "Payroll Export Value": "",
                "Difference": "",
                "Notes": "Employee exists in Workday file but not in payroll export file.",
            })
            continue

        if row["_merge"] == "right_only":
            rows.append({
                "Employee ID": uid,
                "Employee Name": display_name,
                "Mismatch Type": "Missing employee",
                "Field": "Employee ID",
                "Workday Value": "",
                "Payroll Export Value": uid,
                "Difference": "",
                "Notes": "Employee exists in payroll export file but not in Workday file.",
            })
            continue

        for field in compare_fields:
            left = row.get(f"{field} - Workday", 0)
            right = row.get(f"{field} - Payroll Export", 0)
            left_num = 0 if pd.isna(left) else float(left)
            right_num = 0 if pd.isna(right) else float(right)
            diff = round(left_num - right_num, 2)

            if abs(diff) > 0:
                rows.append({
                    "Employee ID": uid,
                    "Employee Name": display_name,
                    "Mismatch Type": "Value mismatch",
                    "Field": field,
                    "Workday Value": round(left_num, 2),
                    "Payroll Export Value": round(right_num, 2),
                    "Difference": diff,
                    "Notes": "",
                })

    return pd.DataFrame(rows)


def invalid_id_report(workday_invalid: pd.DataFrame, toast_invalid: pd.DataFrame) -> pd.DataFrame:
    rows = []

    for _, row in workday_invalid.iterrows():
        rows.append({
            "Source File": "Workday",
            "Employee Name": row.get("employee_name", ""),
            "Employee ID Value": row.get("Employee ID", ""),
            "Issue": "Blank or invalid Employee ID; could not align this row by UID.",
        })

    for _, row in toast_invalid.iterrows():
        rows.append({
            "Source File": "Payroll Export",
            "Employee Name": row.get("employee_name", ""),
            "Employee ID Value": row.get("Employee ID", ""),
            "Issue": "Blank or invalid Employee ID; could not align this row by UID.",
        })

    return pd.DataFrame(rows)


def write_report_to_excel(mismatches: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if mismatches.empty:
            pd.DataFrame([{"Result": "No mismatches found."}]).to_excel(writer, index=False, sheet_name="Mismatches")
        else:
            mismatches.to_excel(writer, index=False, sheet_name="Mismatches")

        # Simple formatting
        workbook = writer.book
        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            for col_cells in sheet.columns:
                max_length = 0
                column_letter = col_cells[0].column_letter
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, min(len(value), 40))
                sheet.column_dimensions[column_letter].width = max(12, max_length + 2)

    output.seek(0)
    return output.read()


# -----------------------------
# UI
# -----------------------------

st.title("Payroll Mismatch Finder")
st.write("Upload the two payroll files, click compare, and download the mismatch report.")

col1, col2 = st.columns(2)

with col1:
    file_a = st.file_uploader("Upload Payroll File 1", type=["xlsx", "csv"], key="file_a")

with col2:
    file_b = st.file_uploader("Upload Payroll File 2", type=["xlsx", "csv"], key="file_b")

compare_clicked = st.button("Compare", type="primary", disabled=not (file_a and file_b))

if compare_clicked:
    try:
        df_a, _ = read_excel_or_csv(file_a, ["employee id", "earning", "regular hours"])
        df_b, _ = read_excel_or_csv(file_b, ["employee id", "earning", "regular hours"])

        type_a = detect_file_type(df_a)
        type_b = detect_file_type(df_b)

        if {type_a, type_b} != {"workday", "toast"}:
            st.error("I could not identify one Workday earning register and one payroll export. Make sure you uploaded the correct two files.")
            st.stop()

        workday_df = df_a if type_a == "workday" else df_b
        toast_df = df_a if type_a == "toast" else df_b

        workday_summary, _ = standardize_workday(workday_df)
        toast_summary, _ = standardize_toast(toast_df)

        mismatches = compare_summaries(workday_summary, toast_summary)

        if mismatches.empty:
            st.success("No mismatches found.")
        else:
            st.error(f"{len(mismatches)} mismatches found.")
            st.dataframe(mismatches, use_container_width=True)

            report_bytes = write_report_to_excel(mismatches)
            st.download_button(
                label="Download Mismatches",
                data=report_bytes,
                file_name="Payroll_Mismatches.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as exc:
        st.error("Something went wrong while comparing the files. Make sure both files are the expected payroll exports.")
        with st.expander("Show error details"):
            st.exception(exc)
elif not (file_a and file_b):
    st.info("Upload both files to begin.")
