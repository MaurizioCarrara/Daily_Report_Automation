# ETL KNIME + Excel/VBA Batch Runner

This script demonstrates the **organizational structure of a fully automated data pipeline**.  
Although the KNIME workflow and VBA modules are proprietary and cannot be published for know-how protection reasons, the framework showcases how to **integrate multiple technologies** with **robust error handling** and **reliable scheduling**.

---

## Problem Solved

Daily report generation that previously required:
- ~1 hour of manual work (including cross-checking multiple data sources)
- Human intervention for Excel updates
- High risk of errors due to repetitive manual operations

---

## Technologies Used

- **Batch Script** — process orchestration  
- **VBScript** — Excel automation *(not included in this repo)*  
- **KNIME** — ETL workflow *(not included in this repo)*  
- **Excel VBA** — business logic *(not included in this repo)*  
- **Windows Task Scheduler** — scheduling and unattended execution  

---

## Key Features

- Pre-execution validation and security checks  
- Comprehensive error handling with clear messages  
- Automatic date calculation (current and next month)  
- Dynamic VBA module import  
- File saving with standardized naming conventions  
- Fully autonomous execution — no manual input required  

---

## Functional Overview

- Checks connectivity to the network drive `R:\`
- Verifies **Excel VBA Trust Access** settings  
- Executes a **KNIME** workflow (`.knwf`) from the command line using the `-reset` flag  
- Runs Excel macros silently through **VBScript**  
- Dynamically imports the `DataProcessing.bas` module and triggers `DataComparison.AVVIO`  
- Exports reports (`.xlsx`) with automatic naming based on the current and next month  
- **Automatic Email Generation (via VBA)**:
  - Creates an Outlook message with a dynamic subject (e.g., “Report YYMM”)  
  - **Attachments**: the two reports just generated  
  - **HTML body** containing a summary **KPI table** (totals, variations, notes)  
- Detailed error messages and non-zero exit codes  
- Automatic cleanup of temporary files  

---

## Prerequisites

- **Windows 10/11**
- **KNIME Analytics Platform** installed at:  
  `%USERPROFILE%\AppData\Local\Programs\KNIME\knime.exe`
- **Microsoft Excel** with:
  - *File → Options → Trust Center → Trust Center Settings → Macro Settings*  
  - Enable **“Trust access to the VBA project object model”**
- **Microsoft Outlook** (for email draft or automatic sending)
- Network drive **`R:\`** mapped and accessible  

---

## How It Works (Summary)

1. **Pre-check** — Verifies network drive, Excel trust access, and KNIME installation.  
2. **KNIME Stage** — Executes ETL workflow in batch mode via  
   `-application org.knime.product.KNIME_BATCH_APPLICATION`.  
3. **Date Handling** — Computes `YY/MM/DD` and next month/year using PowerShell.  
4. **Excel/VBA Stage** — For each file (`DataComparison.xlsx`, `DataComparisonNext.xlsx`):
   - Imports `DataProcessing.bas`
   - Runs macro `DataComparison.AVVIO`
   - Saves results as:
     - `YYMM - Daily Report.xlsx`
     - `YY(next)MM(next) - Daily Report.xlsx`
5. **Email Stage (VBA)** — Creates an **Outlook draft** (or sends directly, if configured)
   with the generated files attached and a KPI summary table embedded in the message body.

---

