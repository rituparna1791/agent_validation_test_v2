import os

# ── Databricks connection ────────────────────────────────────────────────────
# No token needed — auth is handled automatically by the Databricks SDK using
# OAuth (Azure AD / Entra ID) when running inside a Databricks App.
DATABRICKS_HOST  = os.environ.get("DATABRICKS_HOST", "")   # e.g. "adb-1234567890.12.azuredatabricks.net"

# ── SQL Warehouse ────────────────────────────────────────────────────────────
SQL_WAREHOUSE_ID = os.environ.get("SQL_WAREHOUSE_ID", "")   # e.g. "abc123def456"

# ── Unity Catalog defaults (user can override from the UI) ──────────────────
DEFAULT_CATALOG = os.environ.get("DEFAULT_CATALOG", "main")
DEFAULT_SCHEMA  = os.environ.get("DEFAULT_SCHEMA",  "default")

# ── LLM (Databricks Foundation Model API — OpenAI-compatible) ───────────────
# Change the model name to whichever endpoint is enabled in your workspace:
#   "databricks-meta-llama-3-3-70b-instruct"
#   "databricks-dbrx-instruct"
#   "databricks-mixtral-8x7b-instruct"
LLM_MODEL_NAME = os.environ.get("LLM_MODEL_NAME", "databricks-meta-llama-3-3-70b-instruct")
LLM_MAX_TOKENS = 2048
LLM_TEMPERATURE = 0.0   # deterministic SQL generation

# ── Column name conventions ──────────────────────────────────────────────────
# Raw table uses ne_name as the site identifier; final tables use sap_id.
SITE_ID_COL_RAW = os.environ.get("SITE_ID_COL_RAW", "ne_name")   # site column in raw table
SITE_ID_COL_NEW = os.environ.get("SITE_ID_COL_NEW", "sap_id")    # site column in final table

# ── Vendor / technology detection ────────────────────────────────────────────
# Table name prefix → vendor + radio technology
VENDOR_PREFIXES: dict[str, dict] = {
    "er_": {"vendor": "Ericsson", "technology": "5G"},
    "nk_": {"vendor": "Nokia",    "technology": "5G"},
    "sm_": {"vendor": "Samsung",  "technology": "4G"},
}

# ── Sampling ──────────────────────────────────────────────────────────────────
# Validation runs on a random sample rather than the full table.
SAMPLE_ID_COUNT   = int(os.environ.get("SAMPLE_ID_COUNT",   "5"))   # sap_ids / cell_names
SAMPLE_DATE_COUNT = int(os.environ.get("SAMPLE_DATE_COUNT", "3"))   # most-recent partition_dates

# ── Validation thresholds ────────────────────────────────────────────────────
NUMERIC_TOLERANCE   = 0.01   # allowed absolute difference for numeric comparisons
MAX_SAMPLE_ROWS     = 10     # bad rows shown per rule in the report
SCHEMA_SAMPLE_ROWS  = 5      # rows fetched to help the LLM understand the data
