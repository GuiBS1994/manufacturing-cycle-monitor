# Manufacturing Cycle Monitor

Automated production cycle monitoring and reporting solution developed for industrial manufacturing environments.

## Overview

This project was created to solve a real operational visibility problem in a plastic injection manufacturing environment.

Production cycle data was being continuously recorded in the ERP database, generating thousands of records daily, but there was no automated intelligence layer to transform that raw data into actionable operational insights.

The solution automates:

- production cycle monitoring
- operational performance analysis
- deviation identification
- automated email reporting
- scheduled execution without manual intervention

---

## Business Problem

Manufacturing operations needed faster visibility into production cycle deviations.

Challenges identified:

- no automatic monitoring of production cycle performance
- excessive manual effort to extract and analyze data
- delayed response to production inefficiencies
- limited visibility for programming, production and leadership teams
- thousands of ERP records generated daily without automated interpretation

---

## Solution

A Python-based monitoring automation was developed to:

- connect directly to SQL production database
- extract operational production cycle data
- compare actual vs expected cycle performance
- classify performance deviations
- generate HTML email dashboards
- automatically distribute reports to stakeholders
- execute on predefined schedules using Windows Task Scheduler

Two monitoring modes were implemented:

### Operational Monitoring
Focused on active production performance snapshots.

Scheduled executions:

- 06:30
- 14:30
- 22:30

### Complete Monitoring
Full production monitoring report with broader operational visibility.

Scheduled execution:

- 07:00

---

## Technical Architecture

ERP SQL Database
↓
Python Data Extraction Layer
↓
Business Rules / Cycle Analysis Engine
↓
HTML Report Generator
↓
Outlook Email Automation
↓
Windows Task Scheduler Automation

---

## Tech Stack

- Python
- SQL
- PyMySQL
- PyWin32
- Outlook COM Automation
- Windows Task Scheduler

---

## Features

- automated SQL data extraction
- production cycle deviation analysis
- business rule classification
- HTML email dashboard generation
- multi-recipient automated delivery
- scheduled unattended execution
- industrial process monitoring

---

## Project Structure

```bash
manufacturing-cycle-monitor/
│
├── SRC/
│   └── monitor_cycles.py
│
├── Screenshots/
│
├── docs/
│
└── README.md
```

---

## Real Business Impact

This automation reduced manual monitoring effort and improved operational visibility by transforming raw ERP production records into actionable daily intelligence.

Expected operational benefits:

- faster deviation detection
- quicker production response
- improved decision-making
- reduced manual reporting workload
- increased operational transparency

---

## Future Improvements

Planned enhancements:

- KPI dashboard integration
- anomaly detection alerts
- historical trend analysis
- Power BI integration
- machine-level performance indicators
- cost per operation monitoring

---

## Disclaimer

Sensitive production, database and organizational information have been removed or anonymized for portfolio purposes.
