# Transaction Risk Testing Tool (VBA)

This project is an Excel-based tool developed in VBA to support risk-based testing of general ledger data. It has been built to simulate internal audit and control review procedures.

It applies automated rules, sampling logic, and dashboards to help auditors or analysts identify high-risk transactions and document findings efficiently.

## Key Features

- **Rule-Based Risk Flagging**:
  - Flags entries based on:
    - High materiality thresholds
    - Suspicious keywords (e.g., “cash”, “consult”)
    - Weekend posting dates
    - Unapproved vendors

- **Sampling Techniques**:
  - Monetary Unit Sampling — focuses on large-value items
  - Random Sampling — ensures broader coverage

- **Structured Output & Dashboard**:
  - Sampled transactions shown with flag reasons
  - Summary statistics (e.g., % flagged, top reasons)
  - Dynamic charts showing flag frequency and risk breakdown

- **Audit Trail Logging**:
  - Tracks tool usage and timestamp for documentation

## 📁 Project Files

| File | Description |
|------|-------------|
| `AuditLogic.bas` | Main audit and risk-checking logic |
| `Sampling.bas`   | Random and monetary unit sampling methods |
| `Dashboard.bas`  | Code to populate summary stats and generate charts |
| `AuditTool.xlsm` | Full Excel workbook with GUI, inputs, and output |

You can open `.bas` files directly in the VBA editor (Alt + F11 → File → Import).

## Tech Stack
- Microsoft Excel (VBA)
- Dynamic named ranges, charts, and tables
- Macro-enabled `.xlsm` workbook

## Use Case

Ideal for simulating or teaching substantive testing, risk-based sampling, or internal audit logic. Can be extended to support:
- Additional risk rules (e.g., duplicate detection)
- Integrated testing dashboards
- Workflow integration with Excel buttons

