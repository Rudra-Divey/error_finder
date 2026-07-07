"""
Core comparison logic for tbl_pnl_dashboard Excel vs MySQL.
Imported by pnl_dashboard_app.py — not meant to be run directly.
"""

import os
import re
import math
from dataclasses import dataclass

import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv

load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_PATH  = r"C:\Users\HP\OneDrive - BIRLA INSTITUTE OF TECHNOLOGY and SCIENCE\Desktop\NSEProject2\tbl_pnl_dashboard_completed.xlsx"
REPORT_PATH = r"C:\Users\HP\OneDrive - BIRLA INSTITUTE OF TECHNOLOGY and SCIENCE\Desktop\NSEProject2\tbl_pnl_dashboard_report.xlsx"
SHEET_NAME  = "tbl_pnl_dashboard"
TABLE_NAME  = "tbl_pnl_dashboard"
TOLERANCE   = 1e-4
DATE_PATTERN = r"[A-Za-z]+\s+\d{1,2},\s*\d{4}"

FREQUENCY_MAP = {
    "day": "DAY", "week": "WEEK", "month": "MONTH",
    "quarter": "QUARTER", "year": "YEAR",
}

METRIC_ROW_MAP = {
    "cash market : turnover": "CM_ADTV",
    "futures : turnover": "FUTURES_ADTV",
    "options : premium turnover": "OPTIONS_ADTV",
    "currency futures : turnover": "CURRENCY_FUTURE",
    "currency options : turnover": "CURRENCY_OPTION",
    "cash markets": "CM_TC",
    "futures": "FUTURE_TC",
    "options": "OPTIONS_TC",
    "currency derivatives": "CURRENCY_TC",
    "others(irf, wdm, commodity, mutual fund)": "OTHERS",
    "total transaction charges": "TOTAL_TRANSACTION_CHARGES",
    "listing services": "LISTING_SERVICES",
    "data centre and connectivity charges": "D_C_C_CHARGES",
    "operating investment income": "O_INV_INCOME",
    "other operating income": "OTHER_O_INCOME",
    "total operating income": "TOTAL_O_INCOME",
    "interest & other investment income": "INTEREST_OTHER_INV_INCOME",
    "subsidiary dividend": "SUBSIDIARY_DIVIDEND",
    "other income": "OTHER_INCOME",
    "total non operting income": "TOTAL_NON_OP_INCOME",
    "total income": "TOTAL_INCOME",
    "empolyee expenses": "EMP_EXP",
    "clearing and settlement charges": "CLEARING_SETTLEMENT_CHARGES",
    "regulatory expenses": "REGULATORY_EXP",
    "sebi settlement fees / penalty / provisions": "SEBI_SETTLEMENT_FEES_PENALTY_PROV",
    "depriciation": "DEPRICIATION",
    "technology expenses": "TECHNOLOGY_EXP",
    "other variable expense - sms & les, license fees for index": "O_VAR_SMS_LES_LICENSE_FEES_INDEX",
    "other expenses": "OTHER_EXPENSES",
    "total expenditure": "TOTAL_EXPENDITURE",
    "profit before tax before ipft and sgf contribution": "PROFIT_B_TAX_B_IPFT_SGF_CONT",
    "less : contribution to ipft": "LESS_CONT_TO_IPFT",
    "less : provision: contribution to core sgf": "LESS_PRO_CONT_TO_CORE_SGF",
    "profit before exceptional items": "PROFIT_BEFORE_EXCEPTIONAL",
    "less exceptional items": "EXCEPTIONAL_ITEM",
    "profit before tax": "PROFIT_BEFORE_TAX",
    "provision for tax": "LESS_PROVISION_FOR_TAX",
    "profit after tax": "PROFIT_AFTER_TAX",
    "expenditure linked to revenue": "EXPENDITURE_LINKED_REV",
    "expenditure linked to revenue|% of operating revenue": "OPERATING_REV_PERC_EXP",
    "balance expenditure": "BALANCE_EXPENDITURE",
    "balance expenditure|% of operating revenue": "OPERATING_REV_PERC_BAL",
    "operating profit": "OPERATING_PROFIT",
    "operating profit margin (%)": "OPERATING_PROFIT_MARGIN",
    "operating ebitda": "OPERATING_EBITDA",
    "operating ebitda margin (%)": "OPERATING_EBITDA_MARGIN",
    "profit before tax margin (%)": "PROFIT_BEFORE_TAX_MARGIN",
    "pat margin (%)": "PAT_MARGIN",
}


# ── DATA CLASSES ──────────────────────────────────────────────────────────────
@dataclass
class ColumnInfo:
    col_idx: int
    flag: str
    frequency: str
    period_date: object


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────
def get_engine():
    host     = os.environ["MYSQL_HOST"]
    user     = os.environ["MYSQL_USER"]
    password = os.environ["MYSQL_PASSWORD"]
    port     = os.environ.get("MYSQL_PORT", "3306")
    url      = f"mysql+mysqlconnector://{user}:{password}@{host}:{port}/records"
    return create_engine(url)


def parse_excel_columns(raw):
    period_row = raw.iloc[5]
    flag_row   = raw.iloc[7]
    columns, current_frequency, current_period_date = [], None, None

    for col_idx in range(raw.shape[1]):
        period_text = period_row[col_idx]
        if isinstance(period_text, str) and period_text.strip():
            text_lower = period_text.lower()
            for key, freq in FREQUENCY_MAP.items():
                if key in text_lower:
                    current_frequency = freq
                    break
            date_found = re.search(DATE_PATTERN, period_text)
            if date_found:
                current_period_date = pd.to_datetime(date_found.group()).date()

        flag_text = flag_row[col_idx]
        if isinstance(flag_text, str) and flag_text.strip().upper() in ("NSE", "BSE"):
            if current_frequency and current_period_date:
                columns.append(ColumnInfo(
                    col_idx=col_idx,
                    flag=flag_text.strip().upper(),
                    frequency=current_frequency,
                    period_date=current_period_date,
                ))
    return columns


def parse_excel_metric_rows(raw):
    row_map, previous_label = {}, None
    for row_idx in range(raw.shape[0]):
        label = raw.iloc[row_idx, 2]
        if isinstance(label, str):
            clean = label.strip().lower()
            if clean == "% of operating revenue":
                key = f"{previous_label}|% of operating revenue"
            else:
                key = clean
                previous_label = clean
            if key in METRIC_ROW_MAP:
                row_map[row_idx] = METRIC_ROW_MAP[key]
    return row_map


def load_sql_table(engine, columns, metric_rows):
    sql_metric_cols = list(set(metric_rows.values()))
    select_cols     = ", ".join(["DUPLOAD_DATE", "FLAG", "FREQUENCY"] + sql_metric_cols)
    conditions      = [
        f"(DUPLOAD_DATE='{c.period_date}' AND FLAG='{c.flag}' AND FREQUENCY='{c.frequency}')"
        for c in columns
    ]
    query = f"SELECT {select_cols} FROM {TABLE_NAME} WHERE {' OR '.join(conditions)}"
    df    = pd.read_sql(query, engine)
    df["FLAG"]         = df["FLAG"].astype(str).str.upper()
    df["FREQUENCY"]    = df["FREQUENCY"].astype(str).str.upper()
    df["DUPLOAD_DATE"] = pd.to_datetime(df["DUPLOAD_DATE"]).dt.date
    return df


def values_match(excel_val, sql_val):
    if excel_val is None or (isinstance(excel_val, float) and math.isnan(excel_val)):
        return sql_val is None
    if sql_val is None:
        return False
    try:
        return math.isclose(float(excel_val), float(sql_val), abs_tol=TOLERANCE)
    except (TypeError, ValueError):
        return str(excel_val).strip() == str(sql_val).strip()


def run_comparison():
    engine      = get_engine()
    raw         = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)
    columns     = parse_excel_columns(raw)
    metric_rows = parse_excel_metric_rows(raw)

    if not columns:
        raise ValueError("No NSE/BSE columns detected — check sheet layout.")
    if not metric_rows:
        raise ValueError("No metric rows detected in column C — check sheet layout.")

    sql_df = load_sql_table(engine, columns, metric_rows)
    issues = []
    total  = 0

    for row_idx, sql_col in metric_rows.items():
        for col in columns:
            total    += 1
            excel_val = raw.iloc[row_idx, col.col_idx]
            match     = sql_df[
                (sql_df["FLAG"]         == col.flag)       &
                (sql_df["FREQUENCY"]    == col.frequency)  &
                (sql_df["DUPLOAD_DATE"] == col.period_date)
            ]

            if match.empty:
                issues.append({
                    "Row": row_idx + 1, "Col": col.col_idx + 1,
                    "Metric": sql_col, "Flag": col.flag,
                    "Frequency": col.frequency, "Date": col.period_date,
                    "Excel Value": excel_val, "SQL Value": None,
                    "Status": "NO SQL ROW FOUND",
                })
            elif len(match) > 1:
                issues.append({
                    "Row": row_idx + 1, "Col": col.col_idx + 1,
                    "Metric": sql_col, "Flag": col.flag,
                    "Frequency": col.frequency, "Date": col.period_date,
                    "Excel Value": excel_val, "SQL Value": None,
                    "Status": "MULTIPLE ROWS (ambiguous)",
                })
            else:
                sql_val = match.iloc[0][sql_col]
                if not values_match(excel_val, sql_val):
                    issues.append({
                        "Row": row_idx + 1, "Col": col.col_idx + 1,
                        "Metric": sql_col, "Flag": col.flag,
                        "Frequency": col.frequency, "Date": col.period_date,
                        "Excel Value": excel_val, "SQL Value": sql_val,
                        "Status": "MISMATCH",
                    })

    issues_df = pd.DataFrame(issues) if issues else pd.DataFrame()
    return total, issues_df
