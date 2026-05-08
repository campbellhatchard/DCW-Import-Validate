# DCW Import & Validation Tool

This repository contains an implementation of the DCW Import & Validation Tool described in the design specification.  It allows a user to upload a multi‑tab Excel workbook (such as the provided **Final COX DCW.xlsx**), review its contents in a browser, validate the data against business rules, view a consolidated error report and, if there are no errors, trigger the import of the data into staging tables.

The tool is built as a Python web application using [Streamlit](https://streamlit.io/).  Streamlit provides a simple way to create data‑driven web interfaces without writing custom JavaScript.  The backend logic for parsing Excel files and validating the data resides in `validation.py`.  Placeholder import functions live in `logic_blocks.py`; these should be replaced with the real integration logic for your environment.

## Features

* **File upload & review:** Upload `.xlsx` files and view each sheet in a tabbed grid.
* **Validation engine:** Apply the full set of rules defined in the DCW import specification, including mandatory fields, uniqueness constraints, cross‑tab references and enumerated value checks.
* **Error reporting:** Consolidate all validation errors into a single table, filter by tab name and export to CSV.
* **Import trigger:** After a clean validation, click a button to run import logic for each sheet.

## Running locally

### Prerequisites

* Python 3.9 or later

### Installation

Clone this repository and install dependencies:

```bash
git clone <your fork of this repo>
cd dcw_import_app
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Usage

Run the Streamlit application locally:

```bash
streamlit run main.py
```

Navigate to the URL displayed in the console (typically <http://localhost:8501>).  Use the file uploader to select your Excel file and follow the on‑screen instructions.

## Deployment on Render

This repository is configured for deployment on [Render](https://render.com/).  Render automatically builds and runs your application based on the `render.yaml` file in the project root.

To deploy:

1. Push your code to a public or private Git repository (e.g., GitHub).
2. Log into Render and click **New + → Web Service**.
3. Connect your repository and select the branch containing this app.
4. Render will detect the `render.yaml` blueprint and use it to build and run the service.  It will install the dependencies from `requirements.txt` and start Streamlit with the command specified in the blueprint.

Your application will then be accessible via the URL provided by Render.

## Customisation

The validation logic can be extended or modified by editing `validation.py`.  The current implementation covers the rules outlined in the specification, but it may need adjustment if your Excel format changes.  Similarly, `logic_blocks.py` contains placeholder functions that should be replaced with real database or API calls when integrating into a production system.