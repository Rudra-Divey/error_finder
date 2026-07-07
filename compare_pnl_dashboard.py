"""
Compare cell values in tbl_pnl_dashboard_completed.xlsx against rows in the
MySQL table records.tbl_pnl_dashboard.

The Excel sheet is a pivoted dashboard:
  - Row labels (col C, rows 9-11) = metrics -> SQL columns
  - Column headers (rows 5 & 7) = period (Day/Week/Month/Quarter) + exchange (NSE/BSE)
Each individual cell should equal the corresponding SQL column value for the
row in tbl_pnl_dashboard matching that (FLAG, FREQUENCY).

DB credentials are read from environment variables:
  MYSQL_HOST, MYSQL_USER, MYSQL_PASSWORD, MYSQL_DB  (MYSQL_PORT optional, default 3306)
"""

import os
import re
import math
import sys
from dataclasses import dataclass

import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv  

# This loads the environment variables from your .env file
load_dotenv() 

EXCEL_PATH = r"C:\Users\HP\OneDrive - BIRLA INSTITUTE OF TECHNOLOGY and SCIENCE\Desktop\NSEProject2\tbl_pnl_dashboard_completed.xlsx"
SHEET_NAME = "tbl_pnl_dashboard"
TABLE_NAME = "tbl_pnl_dashboard"

# Map the text found in the "For the X ended ..." header to the SQL ENUM value
FREQUENCY_MAP = {
    "day": "DAY",
    "week": "WEEK",
    "month": "MONTH",
    "quarter": "QUARTER",
    "year": "YEAR",
}

# Map Excel row label (lowercased, stripped) -> SQL column name
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
    "other variable expense - sms & les, license fees for index":"O_VAR_SMS_LES_LICENSE_FEES_INDEX",
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
    "expenditure linked to revenue":"EXPENDITURE_LINKED_REV",
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

TOLERANCE = 1e-4  # decimal comparison tolerance


DATE_PATTERN = r"[A-Za-z]+\s+\d{1,2},\s*\d{4}"


@dataclass
class ColumnInfo:
    col_idx: int
    flag: str
    frequency: str
    period_date: object  # the date parsed from the header text (pandas Timestamp)


def get_engine():
    host = os.environ["MYSQL_HOST"]
    user = os.environ["MYSQL_USER"]
    password = os.environ["MYSQL_PASSWORD"]
    db = "records"
    port = os.environ.get("MYSQL_PORT", "3306")
    url = f"mysql+mysqlconnector://{user}:{password}@{host}:{port}/{db}"
    return create_engine(url)


def load_sql_table(engine, columns, metric_rows) -> pd.DataFrame:
    # Work out exactly which SQL columns we need: the key columns + only the metric columns
    # that actually appear in the Excel sheet
    sql_metric_cols = list(set(metric_rows.values()))   # e.g. ["CM_ADTV", "FUTURES_ADTV", "OPTIONS_ADTV"]
    select_cols = ", ".join(["DUPLOAD_DATE", "FLAG", "FREQUENCY"] + sql_metric_cols)

    # Build a WHERE clause that only fetches the exact (date, flag, frequency) rows we need
    # e.g. (DUPLOAD_DATE='2026-01-01' AND FLAG='NSE' AND FREQUENCY='DAY') OR (...)
    conditions = []
    for col in columns:
        condition = (
            f"(DUPLOAD_DATE = '{col.period_date}' "
            f"AND FLAG = '{col.flag}' "
            f"AND FREQUENCY = '{col.frequency}')"
        )
        conditions.append(condition)
    where_clause = " OR ".join(conditions)

    query = f"SELECT {select_cols} FROM {TABLE_NAME} WHERE {where_clause}"

    df = pd.read_sql(query, engine)
    df["FLAG"] = df["FLAG"].astype(str).str.upper()
    df["FREQUENCY"] = df["FREQUENCY"].astype(str).str.upper()
    df["DUPLOAD_DATE"] = pd.to_datetime(df["DUPLOAD_DATE"]).dt.date
    return df


def parse_excel_columns(raw: pd.DataFrame) -> list[ColumnInfo]:
    """Walk the period-header row and the NSE/BSE row to build column metadata."""
    period_row = raw.iloc[5]   # "For the day ended ..." style labels
    flag_row = raw.iloc[7]     # NSE / BSE labels

    columns = []
    current_frequency = None
    current_period_date = None
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
            if current_frequency is None or current_period_date is None:
                raise ValueError(f"Found flag column at idx {col_idx} with no frequency/date context")
            columns.append(ColumnInfo(col_idx=col_idx, flag=flag_text.strip().upper(),
                                       frequency=current_frequency, period_date=current_period_date))
    return columns


def parse_excel_metric_rows(raw: pd.DataFrame) -> dict[int, str]:
    """Map row index -> SQL column name, based on the label in column C (idx 2)."""
    row_map = {}
    previous_label = None                        # memory variable
    for row_idx in range(raw.shape[0]):
        label = raw.iloc[row_idx, 2]
        if isinstance(label, str):
            clean = label.strip().lower()
            if clean == "% of operating revenue":
                key = f"{previous_label}|% of operating revenue"   # combine with row above
            else:
                key = clean
                previous_label = clean           # update memory for next row
            if key in METRIC_ROW_MAP:
                row_map[row_idx] = METRIC_ROW_MAP[key]
    return row_map


def values_match(excel_val, sql_val) -> bool:
    if excel_val is None or (isinstance(excel_val, float) and math.isnan(excel_val)):
        return sql_val is None
    if sql_val is None:
        return False
    try:
        return math.isclose(float(excel_val), float(sql_val), abs_tol=TOLERANCE)
    except (TypeError, ValueError):
        return str(excel_val).strip() == str(sql_val).strip()


def main():
    engine = get_engine()

    # Step 1: parse Excel first so we know exactly what to fetch from SQL
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)

    columns = parse_excel_columns(raw)
    metric_rows = parse_excel_metric_rows(raw)

    if not columns:
        sys.exit("No NSE/BSE columns detected in the Excel header rows - check sheet layout.")
    if not metric_rows:
        sys.exit("No metric rows detected in column C - check sheet layout.")

    # Step 2: fetch only the rows and columns we actually need from SQL
    sql_df = load_sql_table(engine, columns, metric_rows)

    results = []
    for row_idx, sql_col in metric_rows.items():
        for col in columns:
            excel_val = raw.iloc[row_idx, col.col_idx]

            match = sql_df[(sql_df["FLAG"] == col.flag) &
                           (sql_df["FREQUENCY"] == col.frequency) &
                           (sql_df["DUPLOAD_DATE"] == col.period_date)]

            if match.empty:
                results.append({
                    "row": row_idx + 1, "col": col.col_idx + 1,
                    "metric": sql_col, "flag": col.flag, "frequency": col.frequency,
                    "date": col.period_date,
                    "excel_value": excel_val, "sql_value": None,
                    "status": "NO SQL ROW FOUND",
                })
                continue
            if len(match) > 1:
                results.append({
                    "row": row_idx + 1, "col": col.col_idx + 1,
                    "metric": sql_col, "flag": col.flag, "frequency": col.frequency,
                    "date": col.period_date,
                    "excel_value": excel_val, "sql_value": None,
                    "status": "MULTIPLE SQL ROWS MATCHED (ambiguous)",
                })
                continue

            sql_val = match.iloc[0][sql_col]
            ok = values_match(excel_val, sql_val)
            results.append({
                "row": row_idx + 1, "col": col.col_idx + 1,
                "metric": sql_col, "flag": col.flag, "frequency": col.frequency,
                "date": col.period_date,
                "excel_value": excel_val, "sql_value": sql_val,
                "status": "MATCH" if ok else "MISMATCH",
            })

    report = pd.DataFrame(results)

    mismatches = report[report["status"] != "MATCH"]
    print(f"Total cells checked: {len(report)}")
    print(f"Matches: {(report['status'] == 'MATCH').sum()}")
    print(f"Issues:  {len(mismatches)}")
    if not mismatches.empty:
        print("\n--- Issues found ---")
        print(mismatches.to_string(index=False))
    else:
        print("\nAll cells matched their corresponding SQL rows.")

    out_path = r"C:\Users\HP\OneDrive - BIRLA INSTITUTE OF TECHNOLOGY and SCIENCE\Desktop\NSEProject2\tbl_pnl_dashboard_report.xlsx"
    report.to_excel(out_path, index=False)
    print(f"\nFull report written to {out_path}")


if __name__ == "__main__":
    main()
