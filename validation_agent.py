"""
Core validation agent — no LangChain dependency, just direct calls to:
  • Databricks SQL Connector  (schema reads + query execution)
  • Databricks Foundation Model API via openai SDK  (SQL generation)
  • openpyxl  (Excel rule parsing)
"""

from __future__ import annotations

import json
import re
import textwrap
from dataclasses import dataclass, field
from typing import Any, Callable

import openpyxl
from databricks import sql as dbsql
from databricks.sdk import WorkspaceClient
from openai import OpenAI

import config


# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class RuleResult:
    rule_text: str
    columns_in_scope: list[str]   # counter column names extracted from the formula
    aggregation: str              # SUM | AVG | MAX | RATIO
    granularity: str              # daily_cell | daily_site | hourly_cell | hourly_site
    generated_sql: str
    status: str                   # "PASS" | "FAIL" | "ERROR"
    failing_row_count: int = 0
    sample_bad_rows: list[dict] = field(default_factory=list)
    error_message: str = ""


@dataclass
class ValidationReport:
    raw_table: str
    new_table: str
    vendor: str                   # Ericsson | Nokia | Samsung | Unknown
    technology: str               # 5G | 4G | Unknown
    granularity: str              # daily_cell | daily_site | hourly_cell | hourly_site
    total_rules: int
    passed: int
    failed: int
    errors: int
    quality_score: float          # 0–100
    rule_results: list[RuleResult] = field(default_factory=list)
    raw_row_count: int = 0
    new_row_count: int = 0
    sampled_ids: list[str] = field(default_factory=list)   # sap_ids / cell_names used
    sampled_dates: list[str] = field(default_factory=list)
    insights: dict = field(default_factory=dict)           # coverage, nulls, distribution, etc.


# ── Agent ─────────────────────────────────────────────────────────────────────

class ValidationAgent:
    """
    Usage:
        agent = ValidationAgent(warehouse_id="abc123")
        report = agent.run(
            excel_path="/Volumes/main/default/uploads/rules.xlsx",
            raw_table="main.default.raw_orders",
            new_table="main.default.new_orders",
            on_progress=lambda msg: print(msg),
        )

    Auth is handled automatically by the Databricks SDK (OAuth / Azure AD).
    No token needs to be passed — works even when PATs are disabled.
    """

    def __init__(self, warehouse_id: str = config.SQL_WAREHOUSE_ID):
        self.warehouse_id = warehouse_id

        # SDK auto-detects auth: OAuth M2M, Azure AD, or env-var credentials.
        # When running as a Databricks App this requires zero configuration.
        self._sdk = WorkspaceClient()
        # Strip scheme — SQL connector and OpenAI base_url need bare hostname
        self.host = re.sub(r"^https?://", "", self._sdk.config.host.rstrip("/"))

        # Obtain a short-lived OAuth bearer token for the OpenAI-compatible
        # Foundation Model API endpoint.  The SDK refreshes it automatically.
        oauth_token = self._sdk.config.authenticate()          # returns {"Authorization": "Bearer <token>"}
        bearer = oauth_token.get("Authorization", "").replace("Bearer ", "")

        self._llm = OpenAI(
            api_key=bearer,
            base_url=f"https://{self.host}/serving-endpoints",
        )

    # ── Public entry point ────────────────────────────────────────────────────

    def run(
        self,
        excel_path: str,
        raw_table: str,
        new_table: str,
        on_progress: Callable[[str], None] | None = None,
    ) -> ValidationReport:
        log = on_progress or (lambda m: None)

        # ── Vendor / technology ───────────────────────────────────────────────
        vendor_info = self._detect_vendor(new_table)
        log(f"🏭 Vendor: {vendor_info['vendor']}  |  Technology: {vendor_info['technology']}")

        log("📂 Reading validation rules from Excel…")
        rules = self._parse_excel(excel_path)
        log(f"   Found {len(rules)} rule(s)")

        log(f"🔍 Fetching schema for raw table `{raw_table}`…")
        raw_schema = self._get_schema(raw_table)
        log(f"   {len(raw_schema['columns'])} columns, {raw_schema['row_count']:,} rows")

        log(f"🔍 Fetching schema for new table `{new_table}`…")
        new_schema = self._get_schema(new_table)
        log(f"   {len(new_schema['columns'])} columns, {new_schema['row_count']:,} rows")

        gran = self._detect_granularity(new_schema)
        log(f"📏 Detected granularity: {gran['granularity']} — GROUP BY: {', '.join(gran['group_by_raw'])}")

        # ── Sample keys (consistent across all KPI checks) ────────────────────
        log(f"🎲 Sampling {config.SAMPLE_ID_COUNT} IDs × {config.SAMPLE_DATE_COUNT} dates…")
        sample_keys = self._get_sample_keys(new_table, gran)
        log(f"   IDs: {sample_keys['id_values']}")
        log(f"   Dates: {sample_keys['dates']}")

        # ── Insight checks ────────────────────────────────────────────────────
        log("\n🔬 Running insight checks…")
        insights = self._run_insights(raw_table, new_table, gran, sample_keys, new_schema)
        for k, v in insights.items():
            log(f"   {k}: {v}")

        # ── KPI rule validation ───────────────────────────────────────────────
        results: list[RuleResult] = []
        for i, (kpi_name, formula, aggregation) in enumerate(rules, 1):
            counters = self._extract_counters(formula)
            log(f"\n📐 KPI {i}/{len(rules)}: `{kpi_name}` [{aggregation}]")
            log(f"   Formula: {formula}")
            if counters:
                log(f"   Counters detected: {', '.join(counters)}")
            result = self._validate_rule(
                kpi_name, formula, aggregation, gran, sample_keys,
                raw_table, raw_schema,
                new_table, new_schema,
                log,
            )
            results.append(result)
            icon = "✅" if result.status == "PASS" else ("❌" if result.status == "FAIL" else "⚠️")
            log(f"   {icon} {result.status}"
                + (f" — {result.failing_row_count:,} failing rows" if result.status == "FAIL" else ""))

        passed = sum(1 for r in results if r.status == "PASS")
        failed = sum(1 for r in results if r.status == "FAIL")
        errors = sum(1 for r in results if r.status == "ERROR")
        score  = round(100 * passed / len(results), 1) if results else 0.0

        log(f"\n🏁 Done — Quality score: {score}%  ({passed} passed, {failed} failed, {errors} errors)")

        return ValidationReport(
            raw_table=raw_table,
            new_table=new_table,
            vendor=vendor_info["vendor"],
            technology=vendor_info["technology"],
            granularity=gran["granularity"],
            total_rules=len(results),
            passed=passed,
            failed=failed,
            errors=errors,
            quality_score=score,
            rule_results=results,
            raw_row_count=raw_schema["row_count"],
            new_row_count=new_schema["row_count"],
            sampled_ids=sample_keys["id_values"],
            sampled_dates=sample_keys["dates"],
            insights=insights,
        )

    # ── Excel parser ──────────────────────────────────────────────────────────

    def _parse_excel(self, path: str) -> list[tuple[str, str, str]]:
        """
        Returns a list of (kpi_name, formula, aggregation) tuples.

        Expected Excel columns:
          - KPI Name / KPI / Name            — column in the final table  (required)
          - Formula / Calculation / Rule      — CounterFamily.CounterName expression (required)
          - Aggregation / Agg / Agg Type      — SUM | AVG | MAX | RATIO  (default: SUM)

        RATIO means SUM(numerator_counters) / SUM(denominator_counters) — used when the
        formula has two counters divided by each other (e.g. PRB utilisation).
        SUM means SUM(entire counter expression).
        """
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []

        header = [str(c).strip().lower() if c else "" for c in rows[0]]

        kpi_col = None
        for kw in ("kpi name", "kpi", "name", "kpi_name"):
            if kw in header:
                kpi_col = header.index(kw)
                break
        if kpi_col is None:
            kpi_col = 0

        formula_col = None
        for kw in ("formula", "calculation", "rule", "logic"):
            if kw in header:
                formula_col = header.index(kw)
                break
        if formula_col is None:
            formula_col = 1 if len(header) > 1 else 0

        agg_col = None
        for kw in ("aggregation", "agg", "agg type", "agg_type"):
            if kw in header:
                agg_col = header.index(kw)
                break

        results: list[tuple[str, str, str]] = []
        for row in rows[1:]:
            kpi_val     = row[kpi_col]     if kpi_col     < len(row) else None
            formula_val = row[formula_col] if formula_col < len(row) else None
            if not kpi_val or not str(kpi_val).strip():
                continue
            if not formula_val or not str(formula_val).strip():
                continue
            agg_val = "SUM"
            if agg_col is not None and agg_col < len(row) and row[agg_col]:
                agg_val = str(row[agg_col]).strip().upper() or "SUM"
            results.append((str(kpi_val).strip(), str(formula_val).strip(), agg_val))

        return results

    # ── Counter extractor ─────────────────────────────────────────────────────

    @staticmethod
    def _extract_counters(formula: str) -> list[str]:
        """
        Extracts CounterName parts from 'CounterFamily.CounterName' tokens in the formula.
        Example: '(NR MAC SDU Report.DL_MAC_SDU_VOL_DTCH + NR MAC SDU Report.UL_MAC_SDU_VOL_DTCH)'
                 → ['DL_MAC_SDU_VOL_DTCH', 'UL_MAC_SDU_VOL_DTCH']
        """
        matches = re.findall(r'[\w\s]+\.([\w]+)', formula)
        return list(dict.fromkeys(matches))  # deduplicated, order-preserving

    @staticmethod
    def _detect_granularity(new_schema: dict) -> dict:
        """
        Infers granularity from the final table's column set.

        Granularity matrix:
          daily_cell   — cell_name + partition_date, no hour
          daily_site   — sap_id    + partition_date, no hour
          hourly_cell  — cell_name + partition_date + hour
          hourly_site  — sap_id    + partition_date + hour

        Raw table uses ne_name as site id; final table uses sap_id.
        """
        cols_lower = {c["name"].lower() for c in new_schema["columns"]}
        has_hour = "hour" in cols_lower
        is_site  = config.SITE_ID_COL_NEW.lower() in cols_lower

        time_grain = "hourly" if has_hour else "daily"
        level      = "site"   if is_site  else "cell"

        # GROUP BY keys on the raw side (all four ROPs per slot are summed)
        raw_group = [config.SITE_ID_COL_RAW if is_site else "cell_name", "partition_date"]
        new_group = [config.SITE_ID_COL_NEW  if is_site else "cell_name", "partition_date"]
        if has_hour:
            raw_group.append("hour")
            new_group.append("hour")

        # JOIN predicate linking aggregated raw CTE to final table
        id_join = (
            f"N.{config.SITE_ID_COL_NEW} = AGG.{config.SITE_ID_COL_RAW}"
            if is_site else
            "N.cell_name = AGG.cell_name"
        )
        join_predicate = id_join + " AND N.partition_date = AGG.partition_date"
        if has_hour:
            join_predicate += " AND N.hour = AGG.hour"

        return {
            "granularity":     f"{time_grain}_{level}",
            "time_grain":      time_grain,
            "level":           level,
            "is_site":         is_site,
            "has_hour":        has_hour,
            "group_by_raw":    raw_group,
            "group_by_new":    new_group,
            "join_predicate":  join_predicate,
        }

    # ── Vendor detection ──────────────────────────────────────────────────────

    @staticmethod
    def _detect_vendor(table_name: str) -> dict:
        """
        Derives vendor and radio technology from the table name prefix.
          er_  → Ericsson / 5G
          nk_  → Nokia    / 5G
          sm_  → Samsung  / 4G
        Works on fully-qualified names (catalog.schema.table_name).
        """
        bare = table_name.split(".")[-1].lower()
        for prefix, info in config.VENDOR_PREFIXES.items():
            if bare.startswith(prefix):
                return dict(info)
        return {"vendor": "Unknown", "technology": "Unknown"}

    # ── Sample key selector ───────────────────────────────────────────────────

    def _get_sample_keys(self, table_name: str, gran: dict) -> dict:
        """
        Returns a random sample of IDs (sap_id or cell_name) and the most-recent
        partition_dates from the final table.  These are reused across all KPI checks
        so every rule is validated on the same slice of data.
        """
        id_col = config.SITE_ID_COL_NEW if gran["is_site"] else "cell_name"
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT DISTINCT {id_col} FROM {table_name} "
                    f"ORDER BY RAND() LIMIT {config.SAMPLE_ID_COUNT}"
                )
                id_values = [str(r[0]) for r in cur.fetchall()]

                cur.execute(
                    f"SELECT DISTINCT partition_date FROM {table_name} "
                    f"ORDER BY partition_date DESC LIMIT {config.SAMPLE_DATE_COUNT}"
                )
                dates = [str(r[0]) for r in cur.fetchall()]

        return {"id_col": id_col, "id_col_raw": config.SITE_ID_COL_RAW if gran["is_site"] else "cell_name",
                "id_values": id_values, "dates": dates}

    # ── Insight checks ────────────────────────────────────────────────────────

    def _run_insights(
        self,
        raw_table: str,
        new_table: str,
        gran: dict,
        sample_keys: dict,
        new_schema: dict,
    ) -> dict:
        """
        Runs a set of data-quality insight queries independent of formula rules:
          1. Coverage    — % of raw IDs present in the final table
          2. Completeness— expected vs actual row count for the sampled IDs/dates
          3. Null rate   — % nulls across KPI columns in the final table
          4. Distribution— min / max / avg / stddev of each KPI column (sampled slice)
          5. Date gaps   — missing partition_dates in the final table
        """
        id_new = sample_keys["id_col"]
        id_raw = sample_keys["id_col_raw"]
        id_in  = ", ".join(f"'{v}'" for v in sample_keys["id_values"])
        dt_in  = ", ".join(f"'{d}'" for d in sample_keys["dates"])

        # KPI columns = numeric columns that are not dimension/key columns
        dimension_names = {
            "partition_date", "hour", "rop", "cell_name",
            config.SITE_ID_COL_NEW.lower(), config.SITE_ID_COL_RAW.lower(),
        }
        kpi_cols = [
            c["name"] for c in new_schema["columns"]
            if c["name"].lower() not in dimension_names
            and any(t in c["type"].lower() for t in ("int", "long", "float", "double", "decimal"))
        ][:8]  # cap at 8 columns to keep queries short

        insights: dict = {}
        try:
            with self._connect() as conn:
                with conn.cursor() as cur:

                    # 1. Coverage: how many raw IDs made it into the final table
                    cur.execute(f"""
                        SELECT
                            COUNT(DISTINCT R.{id_raw})                          AS raw_ids,
                            COUNT(DISTINCT N.{id_new})                          AS new_ids,
                            ROUND(COUNT(DISTINCT N.{id_new}) * 100.0
                                  / NULLIF(COUNT(DISTINCT R.{id_raw}), 0), 1)   AS coverage_pct
                        FROM {raw_table} R
                        LEFT JOIN {new_table} N ON R.{id_raw} = N.{id_new}
                    """)
                    row = cur.fetchone()
                    insights["coverage"] = {
                        "raw_ids": row[0], "new_ids": row[1],
                        "coverage_pct": float(row[2] or 0),
                    }

                    # 2. Completeness: expected rows vs actual rows for sampled IDs + dates
                    if id_in and dt_in:
                        expected_per_id = 24 if gran["has_hour"] else 1
                        expected_total  = len(sample_keys["id_values"]) * len(sample_keys["dates"]) * expected_per_id
                        cur.execute(f"""
                            SELECT COUNT(*) FROM {new_table}
                            WHERE {id_new} IN ({id_in})
                              AND partition_date IN ({dt_in})
                        """)
                        actual_rows = cur.fetchone()[0]
                        insights["completeness"] = {
                            "expected_rows": expected_total,
                            "actual_rows": actual_rows,
                            "completeness_pct": round(min(actual_rows * 100.0 / max(expected_total, 1), 100), 1),
                        }

                    # 3. Null rate across KPI columns (sampled slice)
                    if kpi_cols and id_in and dt_in:
                        null_exprs = ", ".join(
                            f"ROUND(SUM(CASE WHEN {c} IS NULL THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS {c}"
                            for c in kpi_cols
                        )
                        cur.execute(f"""
                            SELECT {null_exprs}
                            FROM {new_table}
                            WHERE {id_new} IN ({id_in})
                              AND partition_date IN ({dt_in})
                        """)
                        row = cur.fetchone()
                        col_names = [d[0] for d in cur.description]
                        null_rates = {col: float(val or 0) for col, val in zip(col_names, row)}
                        insights["null_rates_pct"] = null_rates

                    # 4. Distribution of KPI columns (sampled slice)
                    if kpi_cols and id_in and dt_in:
                        dist_exprs = ", ".join(
                            f"ROUND(MIN({c}),4) AS {c}_min, ROUND(MAX({c}),4) AS {c}_max, "
                            f"ROUND(AVG({c}),4) AS {c}_avg, ROUND(STDDEV({c}),4) AS {c}_std"
                            for c in kpi_cols
                        )
                        cur.execute(f"""
                            SELECT {dist_exprs}
                            FROM {new_table}
                            WHERE {id_new} IN ({id_in})
                              AND partition_date IN ({dt_in})
                        """)
                        row = cur.fetchone()
                        col_names = [d[0] for d in cur.description]
                        raw_dist = dict(zip(col_names, row))
                        # Re-shape into {kpi: {min, max, avg, std}}
                        dist: dict = {}
                        for c in kpi_cols:
                            dist[c] = {
                                "min": raw_dist.get(f"{c}_min"),
                                "max": raw_dist.get(f"{c}_max"),
                                "avg": raw_dist.get(f"{c}_avg"),
                                "std": raw_dist.get(f"{c}_std"),
                            }
                        insights["distribution"] = dist

                    # 5. Date gaps: find missing dates in final table vs a continuous range
                    cur.execute(f"""
                        SELECT MIN(partition_date), MAX(partition_date),
                               COUNT(DISTINCT partition_date) AS present_dates
                        FROM {new_table}
                    """)
                    row = cur.fetchone()
                    if row[0] and row[1]:
                        insights["date_range"] = {
                            "min_date": str(row[0]), "max_date": str(row[1]),
                            "distinct_dates_present": row[2],
                        }

        except Exception as e:
            insights["insight_error"] = str(e)

        return insights

    # ── Schema reader ─────────────────────────────────────────────────────────

    def _get_schema(self, table_name: str) -> dict:
        """Returns {columns: [{name, type}], sample_rows: [...], row_count: int}"""
        with self._connect() as conn:
            with conn.cursor() as cur:
                # Column info
                cur.execute(f"DESCRIBE TABLE {table_name}")
                raw_cols = cur.fetchall()
                columns = [
                    {"name": r[0], "type": r[1]}
                    for r in raw_cols
                    if r[0] and not r[0].startswith("#")
                ]

                # Sample rows
                cur.execute(f"SELECT * FROM {table_name} LIMIT {config.SCHEMA_SAMPLE_ROWS}")
                col_names = [d[0] for d in cur.description]
                sample_rows = [dict(zip(col_names, row)) for row in cur.fetchall()]

                # Row count
                cur.execute(f"SELECT COUNT(*) FROM {table_name}")
                row_count = cur.fetchone()[0]

        return {"columns": columns, "sample_rows": sample_rows, "row_count": row_count}

    # ── Rule validator ────────────────────────────────────────────────────────

    def _validate_rule(
        self,
        kpi_name: str,
        formula: str,
        aggregation: str,
        gran: dict,
        sample_keys: dict,
        raw_table: str,
        raw_schema: dict,
        new_table: str,
        new_schema: dict,
        log: Callable[[str], None],
    ) -> RuleResult:
        counters  = self._extract_counters(formula)
        rule_text = f"{kpi_name} = {formula}"

        log("   🤖 Generating SQL…")
        try:
            sql = self._generate_sql(
                kpi_name, formula, aggregation, counters, gran, sample_keys,
                raw_table, raw_schema, new_table, new_schema,
            )
            log(f"   SQL preview: {sql[:120].replace(chr(10), ' ')}…")
        except Exception as e:
            return RuleResult(rule_text=rule_text, columns_in_scope=counters,
                              aggregation=aggregation, granularity=gran["granularity"],
                              generated_sql="", status="ERROR",
                              error_message=f"LLM error: {e}")

        log("   ▶️  Running SQL on Databricks…")
        try:
            bad_rows = self._run_sql(sql)
        except Exception as e:
            return RuleResult(rule_text=rule_text, columns_in_scope=counters,
                              aggregation=aggregation, granularity=gran["granularity"],
                              generated_sql=sql, status="ERROR",
                              error_message=f"SQL execution error: {e}")

        bad_rows = bad_rows or []
        return RuleResult(
            rule_text=rule_text,
            columns_in_scope=counters,
            aggregation=aggregation,
            granularity=gran["granularity"],
            generated_sql=sql,
            status="PASS" if len(bad_rows) == 0 else "FAIL",
            failing_row_count=len(bad_rows),
            sample_bad_rows=bad_rows[:config.MAX_SAMPLE_ROWS],
        )

    # ── LLM SQL generation ────────────────────────────────────────────────────

    def _generate_sql(
        self,
        kpi_name: str,
        formula: str,
        aggregation: str,
        counters: list[str],
        gran: dict,
        sample_keys: dict,
        raw_table: str,
        raw_schema: dict,
        new_table: str,
        new_schema: dict,
    ) -> str:
        counter_lower = {c.lower() for c in counters}
        raw_counter_cols = [
            c for c in raw_schema["columns"]
            if c["name"].lower() in counter_lower
        ] or raw_schema["columns"]

        group_by_raw_sql = ", ".join(gran["group_by_raw"])
        group_by_new_sql = ", ".join(gran["group_by_new"])
        join_pred        = gran["join_predicate"]

        # Build sample filter clauses to limit validation to the sampled slice
        id_in_raw  = ", ".join(f"'{v}'" for v in sample_keys["id_values"])
        id_in_new  = id_in_raw
        dt_in      = ", ".join(f"'{d}'" for d in sample_keys["dates"])
        sample_filter_raw = (
            f"WHERE {sample_keys['id_col_raw']} IN ({id_in_raw})\n"
            f"              AND partition_date IN ({dt_in})"
            if id_in_raw and dt_in else ""
        )
        sample_filter_new = (
            f"AND N.{sample_keys['id_col']} IN ({id_in_new})\n"
            f"              AND N.partition_date IN ({dt_in})"
            if id_in_new and dt_in else ""
        )

        agg_instructions = {
            "SUM": (
                "Wrap the ENTIRE counter expression in SUM().\n"
                f"Example: SUM(counter_a + counter_b) / 1024"
            ),
            "AVG": (
                "Wrap the ENTIRE counter expression in AVG().\n"
                f"Example: AVG(counter_a)"
            ),
            "MAX": (
                "Wrap the ENTIRE counter expression in MAX().\n"
                f"Example: MAX(counter_a)"
            ),
            "RATIO": (
                "The formula has a division of two counter groups. "
                "Apply SUM() to the NUMERATOR counters and SUM() to the DENOMINATOR counters SEPARATELY "
                "before dividing — do NOT SUM the whole expression.\n"
                "Example: formula = A / B * 100  →  SUM(A) / SUM(B) * 100\n"
                "Example: formula = (A + B) / (C + D)  →  (SUM(A) + SUM(B)) / (SUM(C) + SUM(D))"
            ),
        }.get(aggregation, "Wrap the ENTIRE counter expression in SUM().")

        system = textwrap.dedent(f"""
            You are a data validation SQL expert for Databricks (Delta Lake / Spark SQL).

            CONTEXT
            -------
            Raw table  : 15-minute granularity data, one row per (cell_name, partition_date, hour, rop).
                         Site identifier column in the raw table: {config.SITE_ID_COL_RAW}
            Final table: pre-aggregated KPI table at granularity = {gran['granularity']}
                         Site identifier column in the final table: {config.SITE_ID_COL_NEW}

            FORMULA NOTATION
            ----------------
            The formula uses "CounterFamily.CounterName" notation.
            - CounterFamily is just a label — IGNORE it.
            - CounterName is the actual column name in the RAW table.
            Replace every "CounterFamily.CounterName" with just "CounterName" in your SQL.

            AGGREGATION TYPE: {aggregation}
            {agg_instructions}

            SQL STRUCTURE TO FOLLOW
            -----------------------
            Build the query as:

            WITH AGG AS (
                SELECT {group_by_raw_sql},
                       -- one column per counter used in the formula, each wrapped in the agg function
                FROM {raw_table}
                {sample_filter_raw}
                GROUP BY {group_by_raw_sql}
            )
            SELECT N.{kpi_name}         AS actual,
                   <agg_formula>        AS expected,
                   N.{group_by_new_sql.replace(', ', ', N.')}
            FROM {new_table} N
            JOIN AGG ON {join_pred}
            WHERE ABS(N.{kpi_name} - <agg_formula>) > {config.NUMERIC_TOLERANCE}
              {sample_filter_new}

            STRICT RULES
            ------------
            1. Use Spark SQL syntax only.
            2. The CTE is named AGG. Alias final table as N.
            3. The sample filter (IDs and dates) MUST appear exactly as shown — in the CTE WHERE clause
               AND repeated in the outer WHERE clause.
            4. Return ONLY rows that VIOLATE the rule (ABS difference > {config.NUMERIC_TOLERANCE}).
            5. Return ONLY the SQL query — no explanation, no markdown fences, no comments.
        """).strip()

        user = textwrap.dedent(f"""
            KPI Name    : {kpi_name}
            Formula     : {formula}
            Aggregation : {aggregation}
            Counters    : {', '.join(counters) if counters else '(infer from formula)'}
            Granularity : {gran['granularity']}
            GROUP BY (raw CTE) : {group_by_raw_sql}
            JOIN condition     : {join_pred}
            Sample filter (raw): {sample_filter_raw or '(none)'}
            Sample filter (new): {sample_filter_new or '(none)'}

            RAW TABLE   : {raw_table}
            Counter columns:
            {json.dumps(raw_counter_cols, indent=2)}
            All columns (for reference):
            {json.dumps(raw_schema['columns'], indent=2)}
            Sample rows : {json.dumps(raw_schema['sample_rows'][:3], indent=2, default=str)}

            FINAL TABLE : {new_table}
            All columns :
            {json.dumps(new_schema['columns'], indent=2)}
            Sample rows : {json.dumps(new_schema['sample_rows'][:3], indent=2, default=str)}

            Write the SQL query that returns rows where N.{kpi_name} does not match the formula.
        """).strip()

        response = self._llm.chat.completions.create(
            model=config.LLM_MODEL_NAME,
            messages=[
                {"role": "system", "content": system},
                {"role": "user",   "content": user},
            ],
            max_tokens=config.LLM_MAX_TOKENS,
            temperature=config.LLM_TEMPERATURE,
        )

        sql = response.choices[0].message.content.strip()
        sql = re.sub(r"^```[a-zA-Z]*\n?", "", sql)
        sql = re.sub(r"\n?```$", "", sql)
        return sql.strip()

    # ── SQL runner ────────────────────────────────────────────────────────────

    def _run_sql(self, sql: str) -> list[dict]:
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(sql)
                col_names = [d[0] for d in cur.description]
                rows = cur.fetchall()
        return [dict(zip(col_names, row)) for row in rows]

    # ── DB connection helper ──────────────────────────────────────────────────

    def _connect(self):
        # credentials_provider lets the SDK handle OAuth token refresh automatically.
        return dbsql.connect(
            server_hostname=self.host,
            http_path=f"/sql/1.0/warehouses/{self.warehouse_id}",
            credentials_provider=self._sdk.config.authenticate,
        )
