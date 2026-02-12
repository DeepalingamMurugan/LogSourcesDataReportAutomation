### reports automation

This project is a toolkit built using Streamlit, designed to automate and streamline various data processing tasks, including JIRA data exporting, Rule Category mapping, and QRadar integration reporting.

## Features

### 1. JIRA Data Exporter

- Connects to JIRA API.
- Exports JIRA issues based on JQL queries.
- Allows custom field selection.

### 2. Rule Category Mapper

- Maps rule categories based on rule names or summaries.
- Uses a configurable `patterns.json` file to define regex patterns for different categories (Anomaly, Behavior, Intrusion, Malware, Test).
- Processes Excel files and appends the mapped category.
- Allows adding new patterns dynamically through the UI.

### 3. QRadar Integration Processor

- Automates the filtration and modification of asset and log source data.
- Processes "Asset List", "Log Sources", and "Formula Sheet" Excel files.
- Performs data cleaning (removing inactive/decommissioned assets, stripping whitespace).
- Handles "Telephony" devices separately.
- Updates formulas in the main sheet based on cross-referenced data.
- Generates a final QRadar Integration Report.

## Prerequisites

- Python 3.11 or higher
- pip (Python package installer)

## Installation

### Setup

1.  Clone the repository:

    ```bash
    git clone repo link
    cd LogSourcesDataReportAutomation
    ```

2.  Create and activate a virtual environment (recommended):

    ```bash
    python -m venv .venv
    # Windows
    .venv\Scripts\activate
    # Linux/Mac
    source .venv/bin/activate
    ```

3.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

    The `requirements.txt` file includes:
    - pandas
    - openpyxl
    - streamlit
    - xlwings
    - jira

## Usage

1.  Run the main Streamlit application:

    ```bash
    streamlit run indexStreamApp.py
    ```

    Alternatively, you can use the provided batch script `Z_mainRun.bat` to launch the application.

2.  The application will open in your default web browser.
3.  Use the navigation menu to select the desired tool:
    - **JIRA Data Exporter**: For exporting data from JIRA.
    - **Rule Category Mapper**: For categorizing rules from an Excel file.
    - **QRadar Integration Processesor**: For generating integration reports.

## Project Structure

- `indexStreamApp.py`: Main entry point and navigation hub.
- `jiraAPItestfrontend.py`: Frontend logic for JIRA Data Exporter.
- `JRCfrontend.py`: Frontend logic for Rule Category Mapper.
- `qradint.py`: Core logic for QRadar Integration Processor.
- `patterns.json`: Configuration file for Rule Category Mapper patterns.
- `requirements.txt`: List of Python dependencies.

## Notes

- This is a beta build and is currently under test.
- The application uses temporary files for processing `qradint.py` operations.
- Ensure clear access to the `patterns.json` file for the Rule Category Mapper to function correctly.
