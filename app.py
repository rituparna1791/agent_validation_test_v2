"""
Databricks App — Streamlit UI for the Validation Agent.
Deploy this as a Databricks App (Apps > Create App > Streamlit).
"""

import os
import tempfile
import time
from io import StringIO

import pandas as pd
import streamlit as st

import config
from validation_agent import ValidationAgent, ValidationReport


# ── Page setup ────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Data Validation Agent",
    page_icon="🔍",
    layout="wide",
)

st.title("🔍 Data Validation Agent")
st.caption("Upload your rules, name your tables — the agent does the rest.")

# ── Sidebar: connection settings ──────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Connection Settings")
    st.markdown("Fill these in once. They are stored only in your session.")

    st.info("Authentication is handled automatically via OAuth (Azure AD / Entra ID). No token needed.")

    host = st.text_input(
        "Databricks Host",
        value=config.DATABRICKS_HOST,
        placeholder="adb-1234567890.12.azuredatabricks.net",
        help="Your workspace hostname (no https://). Usually pre-filled when running as a Databricks App.",
    )
    warehouse_id = st.text_input(
        "SQL Warehouse ID",
        value=config.SQL_WAREHOUSE_ID,
        placeholder="abc123def456",
        help="SQL Warehouses → your warehouse → Connection details → HTTP path last segment.",
    )
    st.divider()
    st.markdown("**Sampling**")
    sample_ids = st.number_input(
        "IDs to sample (sap_id / cell_name)",
        min_value=1, max_value=50,
        value=config.SAMPLE_ID_COUNT,
        help="Number of random sap_ids or cell_names used for all KPI checks.",
    )
    sample_dates = st.number_input(
        "Recent dates to sample",
        min_value=1, max_value=30,
        value=config.SAMPLE_DATE_COUNT,
        help="Most-recent partition_dates included in the sample.",
    )
    llm_model = st.selectbox(
        "LLM Model",
        options=[
            "databricks-meta-llama-3-3-70b-instruct",
            "databricks-dbrx-instruct",
            "databricks-mixtral-8x7b-instruct",
        ],
        index=0,
    )
    st.divider()
    st.markdown("**Excel Rules Format**")
    st.markdown(
        "Three columns:\n\n"
        "| KPI Name | Formula | Aggregation |\n"
        "|----------|---------|-------------|\n"
        "| `mac_traffic_total_gb_pm` | `(NR MAC SDU Report.DL_MAC_SDU_VOL_DTCH + NR MAC SDU Report.UL_MAC_SDU_VOL_DTCH) / (1024*1024)` | `SUM` |\n"
        "| `dl_prb_util_pct` | `NR MAC SDU Report.DL_PRB_USED_DL / NR MAC SDU Report.DL_PRB_AVAIL_DL * 100` | `RATIO` |\n"
        "| `active_ue_avg` | `NR MAC SDU Report.ACTIVE_UE_DL` | `AVG` |\n\n"
        "**KPI Name** = column in your final table.\n\n"
        "**Formula** uses `CounterFamily.CounterName` — family name is the raw table label, "
        "`CounterName` is the column in the raw table.\n\n"
        "**Aggregation** — how counters roll up from 15-min raw:\n"
        "- `SUM` — sum the entire expression (default)\n"
        "- `RATIO` — `SUM(numerator) / SUM(denominator)` (for ratio KPIs)\n"
        "- `AVG` — average the expression\n"
        "- `MAX` — max of the expression\n\n"
        "**Granularity is auto-detected** from the final table columns:\n"
        "- `cell_name + partition_date` → daily_cell\n"
        "- `sap_id + partition_date` → daily_site\n"
        "- `cell_name + partition_date + hour` → hourly_cell\n"
        "- `sap_id + partition_date + hour` → hourly_site"
    )

# ── Main form ─────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)
with col1:
    raw_table = st.text_input(
        "Raw / Source Table",
        placeholder="main.default.raw_orders",
        help="Fully qualified name: catalog.schema.table",
    )
with col2:
    new_table = st.text_input(
        "New / Target Table",
        placeholder="main.default.new_orders",
        help="Fully qualified name: catalog.schema.table",
    )

excel_file = st.file_uploader(
    "Upload Rules Excel (.xlsx)",
    type=["xlsx"],
    help="Excel file with one business rule per row.",
)

run_btn = st.button("▶️  Run Validation", type="primary", use_container_width=True)

# ── Validation run ────────────────────────────────────────────────────────────

if run_btn:
    errors: list[str] = []
    if not host:         errors.append("Databricks Host is required.")
    if not warehouse_id: errors.append("SQL Warehouse ID is required.")
    if not raw_table:    errors.append("Raw table name is required.")
    if not new_table:    errors.append("New table name is required.")
    if not excel_file:   errors.append("Please upload a rules Excel file.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    # Override config values with what the user typed in the sidebar
    config.DATABRICKS_HOST    = host
    config.SQL_WAREHOUSE_ID   = warehouse_id
    config.LLM_MODEL_NAME     = llm_model
    config.SAMPLE_ID_COUNT    = int(sample_ids)
    config.SAMPLE_DATE_COUNT  = int(sample_dates)

    # Save uploaded file to a temp location so openpyxl can read it
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(excel_file.getbuffer())
        tmp_path = tmp.name

    # Chat-style progress log
    st.divider()
    st.subheader("📋 Agent Log")
    log_placeholder = st.empty()
    log_lines: list[str] = []

    def append_log(msg: str):
        log_lines.append(msg)
        log_placeholder.markdown("\n\n".join(log_lines))

    report: ValidationReport | None = None
    try:
        agent = ValidationAgent(warehouse_id=warehouse_id)
        report = agent.run(
            excel_path=tmp_path,
            raw_table=raw_table,
            new_table=new_table,
            on_progress=append_log,
        )
    except Exception as exc:
        st.error(f"Agent failed: {exc}")
        st.stop()
    finally:
        os.unlink(tmp_path)

    # ── Report ────────────────────────────────────────────────────────────────

    st.divider()
    st.subheader("📊 Validation Report")

    # Vendor / technology / granularity badges
    tech_color = "🟢" if report.technology == "5G" else "🔵"
    st.markdown(
        f"**Vendor:** {report.vendor} &nbsp;|&nbsp; "
        f"**Technology:** {tech_color} {report.technology} &nbsp;|&nbsp; "
        f"**Granularity:** `{report.granularity}`"
    )

    # Score card
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Quality Score", f"{report.quality_score}%")
    c2.metric("Total Rules",   report.total_rules)
    c3.metric("✅ Passed",     report.passed)
    c4.metric("❌ Failed",     report.failed)
    c5.metric("⚠️ Errors",    report.errors)

    st.caption(
        f"Raw table: **{report.raw_table}** ({report.raw_row_count:,} rows)  |  "
        f"New table: **{report.new_table}** ({report.new_row_count:,} rows)"
    )

    # Sampled slice info
    if report.sampled_ids:
        st.info(
            f"**Validated on sample** — "
            f"IDs: `{'`, `'.join(report.sampled_ids)}`  |  "
            f"Dates: `{'`, `'.join(report.sampled_dates)}`"
        )

    # ── Insights panel ────────────────────────────────────────────────────────
    if report.insights:
        st.subheader("🔬 Data Quality Insights")
        ins = report.insights

        icols = st.columns(3)

        cov = ins.get("coverage", {})
        if cov:
            icols[0].metric(
                "ID Coverage",
                f"{cov.get('coverage_pct', 0):.1f}%",
                help=f"{cov.get('new_ids',0):,} IDs in final table vs {cov.get('raw_ids',0):,} in raw",
            )

        comp = ins.get("completeness", {})
        if comp:
            icols[1].metric(
                "Row Completeness",
                f"{comp.get('completeness_pct', 0):.1f}%",
                help=f"Expected {comp.get('expected_rows',0):,} rows, found {comp.get('actual_rows',0):,}",
            )

        dr = ins.get("date_range", {})
        if dr:
            icols[2].metric(
                "Dates Present",
                str(dr.get("distinct_dates_present", "—")),
                help=f"{dr.get('min_date','?')} → {dr.get('max_date','?')}",
            )

        null_rates = ins.get("null_rates_pct", {})
        if null_rates:
            with st.expander("Null rates per KPI column (sampled slice)", expanded=False):
                null_df_rows = [{"KPI Column": k, "Null %": v} for k, v in null_rates.items()]
                st.dataframe(pd.DataFrame(null_df_rows), use_container_width=True)

        dist = ins.get("distribution", {})
        if dist:
            with st.expander("Value distribution per KPI column (sampled slice)", expanded=False):
                dist_rows = [
                    {"KPI Column": k,
                     "Min": v.get("min"), "Max": v.get("max"),
                     "Avg": v.get("avg"), "Std Dev": v.get("std")}
                    for k, v in dist.items()
                ]
                st.dataframe(pd.DataFrame(dist_rows), use_container_width=True)

        if "insight_error" in ins:
            st.warning(f"Some insight checks failed: {ins['insight_error']}")

    # Per-rule details
    st.subheader("Rule-by-Rule Breakdown")
    for i, r in enumerate(report.rule_results, 1):
        icon  = "✅" if r.status == "PASS" else ("❌" if r.status == "FAIL" else "⚠️")
        label = f"{icon} Rule {i}: `{r.rule_text}`"

        with st.expander(label, expanded=(r.status != "PASS")):
            st.markdown(f"**Status:** {r.status}")
            st.markdown(f"**Granularity:** `{r.granularity}`  |  **Aggregation:** `{r.aggregation}`")

            if r.columns_in_scope:
                st.markdown(f"**Counters:** `{', '.join(r.columns_in_scope)}`")
            else:
                st.markdown("**Counters:** *(inferred by LLM)*")

            if r.error_message:
                st.error(r.error_message)

            if r.generated_sql:
                st.code(r.generated_sql, language="sql")

            if r.status == "FAIL":
                st.markdown(f"**Failing rows:** {r.failing_row_count:,}")
                if r.sample_bad_rows:
                    st.markdown("**Sample bad rows:**")
                    st.dataframe(pd.DataFrame(r.sample_bad_rows), use_container_width=True)

    # Download report as CSV
    csv_rows = []
    for r in report.rule_results:
        csv_rows.append({
            "Vendor": report.vendor,
            "Technology": report.technology,
            "Granularity": r.granularity,
            "KPI": r.rule_text,
            "Aggregation": r.aggregation,
            "Counters": ", ".join(r.columns_in_scope) if r.columns_in_scope else "",
            "Status": r.status,
            "Failing Rows": r.failing_row_count,
            "Sampled IDs": "; ".join(report.sampled_ids),
            "Sampled Dates": "; ".join(report.sampled_dates),
            "Error": r.error_message,
            "Generated SQL": r.generated_sql,
        })
    csv_buf = StringIO()
    pd.DataFrame(csv_rows).to_csv(csv_buf, index=False)
    st.download_button(
        "⬇️  Download Report CSV",
        data=csv_buf.getvalue(),
        file_name="validation_report.csv",
        mime="text/csv",
    )
