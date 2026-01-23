import streamlit as st
import pandas as pd
import duckdb

st.set_page_config(page_title="Deliverability Explorer", layout="wide")
st.title("📨 Excel Deliverability Explorer (No AI)")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.stop()

xls = pd.ExcelFile(uploaded)
sheet = st.selectbox("Select sheet", xls.sheet_names)

df = pd.read_excel(uploaded, sheet_name=sheet)

st.subheader("Preview")
st.dataframe(df.head(50), use_container_width=True)

# DuckDB setup
con = duckdb.connect()
con.register("data", df)

cols = list(df.columns)

def run_and_show(sql: str):
    try:
        out = con.execute(sql).df()
        st.code(sql, language="sql")
        st.success(f"Returned {len(out)} rows")
        st.dataframe(out, use_container_width=True)

        csv = out.to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV", csv, "result.csv", "text/csv")
    except Exception as e:
        st.error(f"Query failed: {e}")

# -------------------------
# Sheet-specific presets
# -------------------------
st.subheader("Presets")

if sheet == "Failure Details":
    # Column mapping (safe even if names change later)
    default_map = {
        "date": "Date Sent",
        "user": "User",
        "domain": "To Domain",
        "reason": "Detail Category",
        "smtp": "SMTP Code",
        "server": "Server",
        "details": "Details"
    }

    def pick_col(label, key):
        preferred = default_map[key]
        options = cols
        idx = options.index(preferred) if preferred in options else 0
        return st.selectbox(label, options, index=idx)

    m1, m2, m3 = st.columns(3)
    with m1:
        col_domain = pick_col("To Domain column", "domain")
        col_reason = pick_col("Failure Reason column", "reason")
    with m2:
        col_user = pick_col("User column", "user")
        col_smtp = pick_col("SMTP Code column", "smtp")
    with m3:
        col_date = pick_col("Date Sent column", "date")
        col_server = pick_col("Server column", "server")

    p1, p2, p3, p4, p5 = st.columns(5)

    with p1:
        if st.button("Top 10 failing domains"):
            sql = f"""
            SELECT "{col_domain}" AS to_domain, COUNT(*) AS failures
            FROM data
            GROUP BY "{col_domain}"
            ORDER BY failures DESC
            LIMIT 10;
            """
            run_and_show(sql)

    with p2:
        if st.button("Top failure reasons"):
            sql = f"""
            SELECT "{col_reason}" AS reason, COUNT(*) AS cnt
            FROM data
            GROUP BY "{col_reason}"
            ORDER BY cnt DESC
            LIMIT 15;
            """
            run_and_show(sql)

    with p3:
        if st.button("Top SMTP codes"):
            sql = f"""
            SELECT "{col_smtp}" AS smtp_code, COUNT(*) AS cnt
            FROM data
            GROUP BY "{col_smtp}"
            ORDER BY cnt DESC
            LIMIT 20;
            """
            run_and_show(sql)

    with p4:
        if st.button("Top users by failures"):
            sql = f"""
            SELECT "{col_user}" AS user, COUNT(*) AS failures
            FROM data
            GROUP BY "{col_user}"
            ORDER BY failures DESC
            LIMIT 20;
            """
            run_and_show(sql)

    with p5:
        if st.button("Failures by day"):
            # DuckDB date handling: cast timestamp/date text safely
            sql = f"""
            SELECT CAST("{col_date}" AS DATE) AS sent_day, COUNT(*) AS failures
            FROM data
            GROUP BY sent_day
            ORDER BY sent_day DESC
            LIMIT 30;
            """
            run_and_show(sql)

    st.caption("Extra preset ideas: failures by Server, failures by From IP Address, keyword search in Details/Subject.")

elif sheet == "Domains w Lower Deliverability":
    # Straight-forward: show lowest success rate
    p1, p2 = st.columns(2)
    with p1:
        if st.button("Lowest success-rate domains"):
            sql = """
            SELECT Domain, Recipients, Delivered, "Success Rate %" AS success_rate
            FROM data
            ORDER BY "Success Rate %" ASC
            LIMIT 25;
            """
            run_and_show(sql)
    with p2:
        if st.button("Highest recipients among low deliverability"):
            sql = """
            SELECT Domain, Recipients, Delivered, "Success Rate %" AS success_rate
            FROM data
            ORDER BY Recipients DESC
            LIMIT 25;
            """
            run_and_show(sql)

elif sheet == "Failure Reasons":
    # This sheet is a pivot-like table with Unnamed columns; we can still display it cleanly:
    st.info("This sheet looks like a pivot table export (Row Labels + Count). Use filters below or just view it.")
    # No special presets; generic tools below will work.

# -------------------------
# Generic Filter Builder
# -------------------------
st.subheader("Custom Filters (generic)")

operators = ["=", "!=", ">", ">=", "<", "<=", "contains", "starts_with", "ends_with", "in (comma separated)"]

n = st.number_input("Number of filters", min_value=0, max_value=10, value=2, step=1)

filters = []
for i in range(int(n)):
    c1, c2, c3 = st.columns([2, 1, 2])
    with c1:
        col = st.selectbox(f"Column #{i+1}", cols, key=f"col_{i}")
    with c2:
        op = st.selectbox(f"Op #{i+1}", operators, key=f"op_{i}")
    with c3:
        val = st.text_input(f"Value #{i+1}", key=f"val_{i}")
    filters.append((col, op, val))

gb1, gb2, gb3 = st.columns(3)
with gb1:
    group_by = st.multiselect("Group by (optional)", cols, default=[])
with gb2:
    metric = st.selectbox("Metric", ["Show rows", "Count rows"], index=0)
with gb3:
    limit = st.number_input("Limit", min_value=1, max_value=100000, value=1000, step=100)

def build_where(filters):
    where_parts = []
    for col, op, val in filters:
        if not val or str(val).strip() == "":
            continue
        col_sql = f'"{col}"'

        if op in ["=", "!=", ">", ">=", "<", "<="]:
            # attempt numeric compare, else string
            try:
                float(val)
                where_parts.append(f"CAST({col_sql} AS DOUBLE) {op} {val}")
            except:
                where_parts.append(f"CAST({col_sql} AS VARCHAR) {op} '{val}'")
        elif op == "contains":
            where_parts.append(f"CAST({col_sql} AS VARCHAR) ILIKE '%{val}%'")
        elif op == "starts_with":
            where_parts.append(f"CAST({col_sql} AS VARCHAR) ILIKE '{val}%'")
        elif op == "ends_with":
            where_parts.append(f"CAST({col_sql} AS VARCHAR) ILIKE '%{val}'")
        elif op == "in (comma separated)":
            items = [x.strip() for x in val.split(",") if x.strip()]
            quoted = ",".join([f"'{x}'" for x in items])
            where_parts.append(f"CAST({col_sql} AS VARCHAR) IN ({quoted})")

    return " AND ".join(where_parts)

where_clause = build_where(filters)
where_sql = f"WHERE {where_clause}" if where_clause else ""

if metric == "Show rows":
    select_sql = "*"
else:
    select_sql = "COUNT(*) AS row_count"

group_sql = ""
if group_by:
    gb_cols = ", ".join([f'"{c}"' for c in group_by])
    group_sql = f"GROUP BY {gb_cols}"

sql_preview = f"""
SELECT {select_sql}
FROM data
{where_sql}
{group_sql}
LIMIT {int(limit)};
"""

st.caption("Generated SQL (transparency)")
st.code(sql_preview, language="sql")

if st.button("Run Custom Filters"):
    run_and_show(sql_preview)
