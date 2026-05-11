# Excel Pivot Automation using VBA + Claude for Excel

## Overview

This project automates the end-to-end creation of Excel Pivot Tables from raw datasets using a combination of VBA automation, Excel Data Model, and Claude for Excel.

The workflow transforms raw worksheet data into structured Excel tables, loads them into the Data Model, generates a dynamic pivot configuration sheet, and automatically creates multiple Pivot Tables with flexible configurations.

---

# Workflow

## [Step 1 — Run `01_tables_datamodel_button`](01_tables_datamodel_button)

This VBA module prepares the workbook for automation.

### Actions Performed
- Converts worksheet datasets into Excel Tables
- Loads all tables into the Excel Data Model
- Creates the `Pivot_Structure` sheet
- Adds a **Pivot Tables** button
- Assigns the pivot creation macro to the button

### Output
- Structured Excel tables
- Data Model configured
- Pivot_Structure sheet generated
- Pivot execution button ready

---

## [Step 2 — Run `02_prompt`](02_prompt)

The Claude for Excel prompt reads the worksheet tables and automatically generates the `Pivot_Structure` configuration.

### Purpose
Define the structure and logic required for dynamic Pivot Table generation.

### Output
A fully populated `Pivot_Structure` sheet containing:
- Row fields
- Column fields
- Value fields
- Aggregation functions
- Filters
- Filter selections

---

## [Step 3 — Run `03_create_multiple_pivots`](03_create_multiple_pivots)

Click the **Pivot Tables** button to generate all Pivot Tables automatically.

### Features
- Multiple Pivot Tables
- Multiple Row fields
- Multiple Column fields
- Multiple Value fields
- Independent aggregation for each value field
- Multiple filters
- Multiple filter selections
- Automatic descending sort on the first value field
- Separate worksheet for every Pivot Table

---

# Project Structure

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
