# Excel Pivot Automation using VBA + Claude for Excel

---

## Table of Contents

- [Overview](#Overview)

---

## Overview

This project automates the end-to-end creation of Excel Pivot Tables from raw datasets using a combination of

- VBA automation
- Claude for Excel

The workflow transforms raw worksheet data into structured Excel tables, loads them into the Data Model, generates a dynamic pivot configuration sheet, and automatically creates multiple Pivot Tables with flexible configurations.

---

## Workflow

### [Step 1 — Run `01_tables_datamodel_button`](01_tables_datamodel_button)

This VBA module prepares the workbook for automation.

#### Actions Performed
- Converts worksheet datasets into Excel Tables
- Loads all tables into the Excel Data Model
- Creates the `Pivot_Structure` sheet
- Adds a **Pivot Tables** button
- Assigns the "Create_Multiple_Pivots" macro to the button

#### Output
- Structured Excel tables
- Data Model configured
- Pivot_Structure sheet generated
- Pivot execution button ready

### [Step 2 — Run `02_prompt`](02_prompt)

The Claude for Excel prompt reads the worksheet tables and automatically generates the `Pivot_Structure` configuration.

#### Purpose
Define the structure and logic required for dynamic Pivot Table generation.

#### Output
A fully populated `Pivot_Structure` sheet containing:
- Row fields
- Column fields
- Value fields
- Aggregation functions
- Filters
- Filter selections

### [Step 3 — Run `03_create_multiple_pivots`](03_create_multiple_pivots)

Click the **Pivot Tables** button to generate all Pivot Tables automatically.

#### Features
- Multiple Pivot Tables
- Multiple Row fields
- Multiple Column fields
- Multiple Value fields
- Independent aggregation for each value field
- Data validation support for aggregation function selection
- Multiple filters
- Multiple filter selections
- Automatic descending sort on the first value field
- Separate worksheet for every Pivot Table

### Supported Aggregation Functions

- Sum
- Count
- Average
- Min
- Max
- DistinctCount

### Key Capabilities

- Dynamic pivot configuration
- Multiple Pivot Table generation
- Excel Data Model integration
- AI-assisted Pivot Structure creation
- Multiple filters support
- Independent aggregation logic
- Automated sorting
- End-to-end reporting automation

---

## Project Structure

```text
01_tables_datamodel_button
│
├── Creates Excel Tables
├── Loads data into Data Model
├── Creates Pivot_Structure sheet
└── Adds Pivot Tables button

02_prompt
│
└── Claude for Excel prompt used to generate Pivot_Structure

03_create_multiple_pivots
│
└── Dynamic Pivot Table generation engine
```

---

## Pivot_Structure Sample

<img width="1201" height="271" alt="Pivot_Structure" src="https://github.com/user-attachments/assets/587bb34e-8902-44a5-bb0b-0968a30bbe10" />
