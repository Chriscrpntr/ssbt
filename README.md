# ssbt — Spreadsheet Build Tool

dbt for engineers who live in Excel.

Turn spreadsheets into a version-controlled, testable build pipeline using SQL and YAML — no database required.

## Quick Start

```
ssbt build          # run models + run tests
ssbt build --dry-run    # show compiled SQL without executing
ssbt test           # run schema tests only
```

### Build output

```
  [raw_orders]
  [completed_orders]
  [enriched_orders]
  [region_summary]
  [top_customers]
  ✓ raw_orders.order_id (not_null)
  ✓ raw_orders.order_id (unique)
  ✓ raw_orders.status (accepted_values)
  ✓ completed_orders.total (not_null)
  ✓ completed_orders.total (positive)
  ✓ enriched_orders.order_id (not_null)
  ✓ enriched_orders.order_id (unique)
  ✓ enriched_orders.email (not_null)

8 passed, 0 failed
Done.
```

### Test output (ssbt test)

```
  ✓ raw_orders.order_id (not_null)
  ✓ raw_orders.order_id (unique)
  ✓ raw_orders.status (accepted_values)
  ✓ completed_orders.total (not_null)
  ✓ completed_orders.total (positive)
  ✓ enriched_orders.order_id (not_null)
  ✓ enriched_orders.order_id (unique)
  ✓ enriched_orders.email (not_null)

8 passed, 0 failed
```

### Test marks

| Mark | Meaning |
|---|---|
| ✓ | Test passed |
| ✗ | Test failed |
| ! | Test skipped/warned (e.g., uniqueness test skipped due to nulls) |

### Failures

```
  ✓ raw_orders.order_id (not_null)
  ✓ raw_orders.order_id (unique)
  ✓ raw_orders.status (accepted_values)
  ✗ completed_orders.total (positive)
  ✓ enriched_orders.order_id (not_null)
  ✓ enriched_orders.order_id (unique)
  ✓ enriched_orders.email (not_null)

7 passed, 1 failed
  FAIL completed_orders.total (positive): 2 non-positive values: [0, -5]
```

```
my-project/
├── ssbt.yml              # manifest: models, dependencies, tests
├── input/                # source spreadsheets
│   ├── orders.xlsx
│   └── customers.xlsx
├── output/               # generated output (created by ssbt)
│   ├── raw_orders.xlsx
│   ├── completed_orders.xlsx
│   └── region_summary.xlsx
├── models/
│   ├── raw_orders.sql
│   ├── completed_orders.sql
│   └── region_summary.sql
└── tests/                # model tests (optional)
    └── test_schema.yml
```

## ssbt.yml

### Top-level keys

| Key | Required | Description |
|---|---|---|
| `name` | yes | Project name |
| `version` | no | Project version string |
| `sources` | no | List of input file paths (see Inputs below) |
| `models` | yes | List of model definitions |

### Sources (multi-file inputs)

```yaml
sources:
  - name: orders
    path: input/orders.xlsx
    sheets:
      - raw_orders
  - name: customers
    path: input/customers.xlsx
    sheets:
      - customer_list
```

Each source registers its sheets as DuckDB tables named `{source_name}_{sheet_name}`.
Models reference them via `{{ ref('orders_raw_orders') }}` or `{{ ref('raw_orders') }}` (if unambiguous).

### Models

```yaml
models:
  - name: completed_orders
    path: models/completed_orders.sql
    config:
      output: output/completed_orders.xlsx
      output_sheet: completed_orders
    columns:
      - name: total
        tests:
          - not_null
          - positive
```

Dependencies are **inferred from `{{ ref('name') }}` in the SQL files** — no `depends_on` needed.

| Key | Description |
|---|---|
| `name` | Model name — used as table name in DuckDB and as the default output sheet name |
| `path` | Path to the SQL file (relative to ssbt.yml) |
| `config.output` | Output file path (default: `{output_dir}/{model_name}.xlsx`) |
| `config.output_sheet` | Output sheet name within the file (default: model name) |
| `columns[].name` | Column name to test |
| `columns[].tests[]` | Test definitions |

By default each model writes to its own Excel file in the output directory. Set `config.output` to control the file path.

### Schema Tests

Tests can be written as a string (no args) or a dict (with args):

```yaml
tests:
  - not_null
  - unique
  - accepted_values:
      values: ["A", "B", "C"]
  - positive
  - not_empty
  - regex_match:
      expression: "^[A-Z]{3}-\\d+$"
```

| Test | Args | Description |
|---|---|---|
| `not_null` | — | Column has no NULL values |
| `unique` | — | Column has no duplicate values |
| `accepted_values` | `values: [...]` | All values are in the allowed set |
| `positive` | — | All values are > 0 |
| `not_empty` | — | No empty string values |
| `regex_match` | `expression: "..."` | All non-null values match the regex |

### Example ssbt.yml

```yaml
name: orders-pipeline
version: "1.0"

sources:
  - name: orders
    path: input/orders.xlsx
    sheets:
      - raw_orders
  - name: customers
    path: input/customers.xlsx
    sheets:
      - customer_list

models:
  - name: raw_orders
    path: models/raw_orders.sql
    config:
      output: output/raw_orders.xlsx
      output_sheet: raw_orders
    columns:
      - name: order_id
        tests:
          - not_null
          - unique

  - name: completed_orders
    path: models/completed_orders.sql
    config:
      output: output/completed_orders.xlsx

  - name: enriched_orders
    path: models/enriched_orders.sql
    config:
      output: output/enriched_orders.xlsx
    columns:
      - name: order_id
        tests:
          - not_null
          - unique
      - name: email
        tests:
          - not_null

  - name: region_summary
    path: models/region_summary.sql
    config:
      output: output/region_summary.xlsx

  - name: top_customers
    path: models/top_customers.sql
    config:
      output: output/top_customers.xlsx
```

## How It Works

1. **Parse** `ssbt.yml` — load models and sources
2. **Infer dependencies** — scan SQL files for `{{ ref('name') }}` calls to build the DAG
3. **Resolve DAG** — topological sort with cycle detection
4. **Compile SQL** — resolve `{{ ref('name') }}` into subqueries
5. **Register sources** — load source sheets from input files into DuckDB tables
6. **Execute** — run models in DAG order, materializing each result in DuckDB
7. **Write output** — write each model's result to its configured output file
8. **Test** — `ssbt test` runs schema tests, writes results to `test_results` sheet

## CLI

```
ssbt build [--yml FILE] [--input FILE] [--output DIR] [--dry-run] [--select MODEL ...]
ssbt test  [--yml FILE] [--input FILE] [--output DIR] [--select MODEL ...]
ssbt docs  [--yml FILE]
```

| Flag | Default | Description |
|---|---|---|
| `--yml` | `ssbt.yml` | Path to manifest |
| `--input` | `input.xlsx` | Default input file (for single-file mode) |
| `--output` | `output/` | Output directory (per-model files) |
| `--dry-run` | — | Show compiled SQL without executing |
| `--select` | — | Only build/test specified models (`ssbt build --select enriched_orders`) |

## Docs

```
ssbt docs
```

Outputs project documentation to stdout:

```
============================================================
  orders-pipeline
  5 models, 2 sources
============================================================

Dependency graph:
----------------------------------------
  raw_orders
  completed_orders <- raw_orders
  enriched_orders
  region_summary <- enriched_orders
  top_customers <- enriched_orders

Sources:
----------------------------------------
  orders:
    file: input.xlsx
    sheet: raw_orders -> table: orders_raw_orders
  customers:
    file: data/customers.xlsx
    sheet: customer_list -> table: customers_customer_list

Models:
----------------------------------------
  raw_orders:
    sql: models/raw_orders.sql
    output: output/raw_orders.xlsx
    columns:
      order_id: not_null
      order_id: unique
      status: accepted_values(...)

  completed_orders:
    sql: models/completed_orders.sql
    output: output/completed_orders.xlsx
    depends_on: raw_orders
    columns:
      total: not_null
      total: positive
...
```

## SQL

Models are plain SQL queries that run against DuckDB. Use `{{ ref('name') }}` to reference other models — it gets resolved into a subquery at compile time.

```sql
-- models/enriched_orders.sql
SELECT
    o.order_id,
    o.customer,
    o.region,
    o.total,
    c.email
FROM {{ ref('completed_orders') }} o
JOIN {{ ref('customers_customer_list') }} c ON o.customer = c.name
WHERE o.status = 'completed'
```

For source sheets, use the `{source_name}_{sheet_name}` table name:

```sql
SELECT * FROM {{ ref('orders_raw_orders') }}
```

## Requirements

- Python 3.9+
- `duckdb`, `openpyxl`, `pandas`, `pyyaml`

```
pip install duckdb openpyxl pandas pyyaml
```
