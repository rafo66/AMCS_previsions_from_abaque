# Exception Report Processing Tool

## Overview

This project processes **Exception Reports** to compute backlog, match productivities, and generate a formatted Excel output.

The script parses the exception report, applies routing corrections, computes forecasts and backlog indicators, and exports the results using a predefined Excel template.

---

# Version

**V1 — 16/03/2026**

---

# Known Issues

### Exception Report Parsing

The latest exception report format is not parsed correctly:

* Some columns are missing (notably **Forecast W5-8**).
* Total sums appear incorrect, possibly due to a filtering issue.

| Report                                                 | Status                  |
| ------------------------------------------------------ | ----------------------- |
| `Exception_report_20260309 - PLEASE DON'T MODIFY.xlsb` | ✅ Works (9/13)          |
| `Exception_report_20260313 - PLEASE DON'T MODIFY.xlsx` | ❌ Does not work (13/13) |

---



### Free FG % above 100% and not relative to next week.


# LASS Bug

Issue with totals and details caused by incorrect routing values.

### Fix implemented

Inside `filterFalseBacklog`:

```python
# if LAS1 in Routing, replace by LAS1.
self.exceptionReportDF["Routing"] = self.exceptionReportDF.apply(
    lambda row: "LAS1." if "LAS1" in str(row["Routing"]) else row["Routing"],
    axis=1
)
```

---

# Processing Rules

## 1. Backlog Detection

A row is delated from backlog if:

```
(Hire Work) & (Backlog > 20) & (forecast_sum == 0)
```

---

## 2. Prediction Adjustments (New Abaque)

Prediction adjustments relative to real productivity:

| Routing | Adjustment |
| ------- | ---------- |
| L1      | +20%       |
| R1 & R6 | +5%        |
| P3      | -20%       |
| LAS     | -35%       |

---

# Project Structure (Planned Refactor)

The code should be split into **three scripts**:

### `main.py`

Contains:

* `excelHandler` class
* `MatchingProductivities` class
* Core logic to:

  * Process exception reports
  * Compute productivities

### `matchingProductivities.py`

Contains:

* `MatchingProductivities` class
* Logic for matching articles with productivity data.

### `outputFormatter.py`

Contains:

* `outputFormatter` class
* Excel formatting logic for the output file.

---

# TODO

### Functional Improvements

* [ ] Add `clear_last` feature

  * Clear **details table**
  * Clear **global summary table**



---

# Dependencies

The project relies on the following files:

```
Reports/Report.xlsb
Abaque/Abaque.xlsm
OutputTemplate.xlsm
```

### Cached Files

Generated automatically:

```
Processed_Exception_report.xlsx
articleCachedProductivities.txt
```

---

# Environment Setup

Activate the virtual environment:

```bash
source .env/bin/activate
```

---

# Automation (Linux Cron)

Edit cron jobs:

```bash
crontab -e
```

---

## Run at 6:00 on weekdays

```bash
0 6 * * 1-5 /home/Raftests/AMCS/bots_previsions/semaine_postes/.env/bin/python /home/Raftests/AMCS/bots_previsions/semaine_postes/excelHandler.py >> /home/Raftests/AMCS/bots_previsions/semaine_postes/log.txt 2>&1
```

---

## Run every 10 minutes

```bash
*/10 * * * * /home/Raftests/AMCS/bots_previsions/semaine_postes/.env/bin/python /home/Raftests/AMCS/bots_previsions/semaine_postes/excelHandler.py >> /home/Raftests/AMCS/bots_previsions/semaine_postes/log.txt 2>&1
```

---

## Recommended: Prevent Multiple Instances (using flock)

```bash
30 5 * * * flock -n /tmp/excelHandler.lock /home/Raftests/AMCS/bots_previsions/semaine_postes/.env/bin/python /home/Raftests/AMCS/bots_previsions/semaine_postes/excelHandler.py >> /home/Raftests/AMCS/bots_previsions/semaine_postes/log.txt 2>&1

*/10 * * * * flock -n /tmp/excelHandler.lock /home/Raftests/AMCS/bots_previsions/semaine_postes/.env/bin/python /home/Raftests/AMCS/bots_previsions/semaine_postes/excelHandler.py >> /home/Raftests/AMCS/bots_previsions/semaine_postes/log.txt 2>&1
```

---

# Logs

Monitor execution logs:

```bash
tail -f log.txt
```

---

# Summary

This tool automates the processing of **Exception Reports**, applies routing corrections and productivity adjustments, and generates formatted Excel outputs for operational forecasting.

Planned refactoring will improve maintainability by separating processing, productivity matching, and output formatting into dedicated modules.
