import io
import re
from typing import List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Payroll Comparer", page_icon="📊", layout="wide")

st.title("Payroll Comparer")
st.write("Upload two payroll files, align by Employee UID, and download an Excel file containing only mismatches.")


def normalize_col(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def normalize_value(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    # Normalize numeric-looking values so 40 and 40.0 match
    try:
        num = float(text.replace(",", ""))
        if num.is_integer():
            return str(int(num))
        return str(round(num, 6)).rstrip("0").rstrip(".")
    except Exception:
        return re.sub(r"\s+", " ", text).lower()


def load_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str)
    raise ValueError("Unsupported file type. Please upload CSV or Excel files.")


def likely_uid_columns(columns: List[str]) -> List[str]:
    preferred = []
    for col in columns:
        n = normalize_col(col)
        if n in {"employeeuid", "uid", "employeeid", "empid", "id"} or ("uid" in n) or ("employee" in n and "id" in n):
            preferred.append(col)
    return preferred + [c for c in columns if c not in preferred]


def compare_files(df1: pd.DataFrame, df2: pd.DataFrame, uid1: str, uid2: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df1 = df1.copy()
    df2 = df2.copy()

    df1["__uid__"] = df1[uid1].astype(str).str.strip()
    df2["__uid__"] = df2[uid2].astype(str).str.strip()

    df1 = df1[df1["__uid__"].ne("") & df1["__uid__"].notna()]
    df2 = df2[df2["__uid__"].ne("") & df2["__uid__"].notna()]

    # If duplicate UIDs exist, keep row number information and compare first occurrence.
    dupes = []
    for label, df in [("File 1", df1), ("File 2", df2)]:
        duplicated = df[df["__uid__"].duplicated(keep=False)]
        for uid in sorted(duplicated["__uid__"].unique()):
            dupes.append({"File": label, "Employee UID": uid, "Issue": "Duplicate UID found; first occurrence used for comparison"})

    df1i = df1.drop_duplicates("__uid__", keep="first").set_index("__uid__")
    df2i = df2.drop_duplicates("__uid__", keep="first").set_index("__uid__")

    # Match columns by normalized names, excluding UID columns.
    cols1 = {normalize_col(c): c for c in df1i.columns if c != uid1}
    cols2 = {normalize_col(c): c for c in df2i.columns if c != uid2}
    common_norm_cols = sorted(set(cols1).intersection(cols2))

    rows = []
    all_uids = sorted(set(df1i.index).union(set(df2i.index)))

    for uid in all_uids:
        in1 = uid in df1i.index
        in2 = uid in df2i.index
        if not in1:
            rows.append({"Employee UID": uid, "Field": "ROW", "File 1 Value": "Missing", "File 2 Value": "Present", "Issue": "Employee only in File 2"})
            continue
        if not in2:
            rows.append({"Employee UID": uid, "Field": "ROW", "File 1 Value": "Present", "File 2 Value": "Missing", "Issue": "Employee only in File 1"})
            continue

        for norm_col in common_norm_cols:
            c1 = cols1[norm_col]
            c2 = cols2[norm_col]
            v1 = df1i.at[uid, c1]
            v2 = df2i.at[uid, c2]
            if normalize_value(v1) != normalize_value(v2):
                rows.append({
                    "Employee UID": uid,
                    "Field": c1 if c1 == c2 else f"{c1} / {c2}",
                    "File 1 Value": "" if pd.isna(v1) else v1,
                    "File 2 Value": "" if pd.isna(v2) else v2,
                    "Issue": "Mismatch",
                })

    mismatches = pd.DataFrame(rows)
    duplicate_report = pd.DataFrame(dupes)
    return mismatches, duplicate_report


def to_excel_bytes(mismatches: pd.DataFrame, duplicate_report: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        mismatches.to_excel(writer, index=False, sheet_name="Mismatches")
        if not duplicate_report.empty:
            duplicate_report.to_excel(writer, index=False, sheet_name="Duplicate UIDs")
        workbook = writer.book
        for sheet_name, df in [("Mismatches", mismatches), ("Duplicate UIDs", duplicate_report)]:
            if sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df.columns):
                    width = min(max(len(str(col)), *(len(str(x)) for x in df[col].astype(str).head(100))) + 2, 45)
                    worksheet.set_column(idx, idx, width)
                header_fmt = workbook.add_format({"bold": True})
                for idx, col in enumerate(df.columns):
                    worksheet.write(0, idx, col, header_fmt)
    return output.getvalue()


col_a, col_b = st.columns(2)
with col_a:
    file1 = st.file_uploader("File 1", type=["csv", "xlsx", "xls"], key="file1")
with col_b:
    file2 = st.file_uploader("File 2", type=["csv", "xlsx", "xls"], key="file2")

if file1 and file2:
    try:
        df1 = load_file(file1)
        df2 = load_file(file2)

        st.subheader("Confirm Employee UID Columns")
        col1, col2 = st.columns(2)
        with col1:
            uid_options1 = likely_uid_columns(list(df1.columns))
            uid1 = st.selectbox("Employee UID column in File 1", uid_options1)
            st.caption(f"Rows: {len(df1):,} | Columns: {len(df1.columns):,}")
        with col2:
            uid_options2 = likely_uid_columns(list(df2.columns))
            uid2 = st.selectbox("Employee UID column in File 2", uid_options2)
            st.caption(f"Rows: {len(df2):,} | Columns: {len(df2.columns):,}")

        with st.expander("Preview uploaded files"):
            p1, p2 = st.columns(2)
            with p1:
                st.write("File 1 preview")
                st.dataframe(df1.head(10), use_container_width=True)
            with p2:
                st.write("File 2 preview")
                st.dataframe(df2.head(10), use_container_width=True)

        if st.button("Compare Files", type="primary"):
            mismatches, duplicate_report = compare_files(df1, df2, uid1, uid2)

            if mismatches.empty and duplicate_report.empty:
                st.success("No mismatches found.")
            else:
                st.warning(f"Found {len(mismatches):,} mismatch rows.")
                st.dataframe(mismatches, use_container_width=True)

                excel_bytes = to_excel_bytes(mismatches, duplicate_report)
                st.download_button(
                    "Download Mismatches.xlsx",
                    data=excel_bytes,
                    file_name="Mismatches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if not duplicate_report.empty:
                    st.info("Duplicate Employee UIDs were found. They are included on a separate Excel sheet.")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Upload both files to start.")
