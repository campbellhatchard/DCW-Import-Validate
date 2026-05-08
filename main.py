"""
main.py

This Streamlit application provides a web interface for uploading and
validating DCW import spreadsheets.  Users can upload an Excel file,
review each sheet in a tabbed data grid, run validation against the
business rules defined in validation.py, view a consolidated error
report and trigger an import placeholder when no errors remain.
"""

from __future__ import annotations

import io
from typing import Dict

import pandas as pd
import streamlit as st

from validation import validate_workbook
import logic_blocks


def load_workbook(file_data: bytes) -> Dict[str, pd.DataFrame]:
    """
    Load an Excel workbook into a dictionary of DataFrames keyed by
    sheet name.  This helper reads all sheets at once and returns
    them for validation and display.
    """
    xl = pd.ExcelFile(io.BytesIO(file_data))
    work = {}
    for sheet_name in xl.sheet_names:
        # parse with header=0 (first row as header)
        df = xl.parse(sheet_name, header=0)
        work[sheet_name] = df
    return work


def display_dataframes(tabs, workbook: Dict[str, pd.DataFrame]) -> None:
    """Render each DataFrame in its own tab."""
    for tab, (name, df) in zip(tabs, workbook.items()):
        with tab:
            st.subheader(name)
            # Copy df for display modifications (drop sample rows if needed)
            display_df = df.copy()
            # For Serials and Users sheets drop the first row (sample)
            if name in {"Serials", "Users"} and len(display_df) > 1:
                display_df = display_df.iloc[1:].reset_index(drop=True)
            st.dataframe(display_df, use_container_width=True)


def run_import(workbook: Dict[str, pd.DataFrame]) -> None:
    """
    Call the placeholder import logic for each sheet.  This is run
    after successful validation.  It displays a status for each table.
    """
    st.info("Beginning import...", icon="ℹ️")
    results = {}
    # Map sheet names to import functions defined in logic_blocks
    import_map = {
        "Org Units": logic_blocks.import_org_units,
        "Org Unit Relationships": logic_blocks.import_org_unit_relationships,
        "Customers": logic_blocks.import_customers,
        "Suppliers": logic_blocks.import_suppliers,
        "Zones": logic_blocks.import_zones,
        "ABC Codes": logic_blocks.import_abc_codes,
        "Inventory Locations": logic_blocks.import_inventory_locations,
        "Location Groups": logic_blocks.import_location_groups,
        "Item Global Master": logic_blocks.import_item_global_master,
        "Item Org Units": logic_blocks.import_item_org_units,
        "UOM Conversions": logic_blocks.import_uom_conversions,
        "Item Capacities": logic_blocks.import_item_capacities,
        "Item Tolerances": logic_blocks.import_item_tolerances,
        "Item Cross Ref UPC": logic_blocks.import_item_cross_ref_upc,
        "Item Cross Ref Supplier": logic_blocks.import_item_cross_ref_supplier,
        "Fixed Pick Items": logic_blocks.import_fixed_pick_items,
        "Jobs": logic_blocks.import_jobs,
        "Serials": logic_blocks.import_serials,
        "Inventory Balances": logic_blocks.import_inventory_balances,
        "Users": logic_blocks.import_users,
    }
    for sheet_name, import_fn in import_map.items():
        df = workbook.get(sheet_name)
        if df is None:
            st.warning(f"Sheet '{sheet_name}' missing – skipping import.")
            continue
        # Drop sample row for serials and users before import
        if sheet_name in {"Serials", "Users"} and len(df) > 1:
            df_use = df.iloc[1:].reset_index(drop=True)
        else:
            df_use = df
        try:
            result = import_fn(df_use)
            results[sheet_name] = result
            if result.get("success"):
                st.success(f"Imported {sheet_name}: {result.get('count', 0)} records")
            else:
                st.error(f"Failed to import {sheet_name}: {result}")
        except Exception as exc:
            st.error(f"Exception importing {sheet_name}: {exc}")
    st.info("Import complete.")


def main() -> None:
    st.set_page_config(page_title="DCW Import & Validation Tool", layout="wide")
    st.title("DCW Import & Validation Tool")
    st.write(
        "Upload a multi‑tab Excel workbook (.xlsx), review its contents, \
         validate the data and import it into your staging tables."
    )

    uploaded_file = st.file_uploader(
        "Select Excel file", type=["xlsx", "xls"], help="Choose the DCW import workbook to process"
    )

    # Keep workbook and validation state in session
    if uploaded_file is not None:
        try:
            workbook = load_workbook(uploaded_file.getvalue())
        except Exception as exc:
            st.error(f"Failed to read workbook: {exc}")
            return
        # Display workbook sheets in tabs
        sheet_names = list(workbook.keys())
        tabs = st.tabs(sheet_names)
        display_dataframes(tabs, workbook)

        if st.button("Validate", type="primary"):
            with st.spinner("Validating..."):
                errors = validate_workbook(workbook)
            if errors:
                st.error(f"Validation completed: {len(errors)} errors found.")
                err_df = pd.DataFrame(errors)
                st.dataframe(err_df, use_container_width=True)
                # Provide download of errors as CSV
                csv = err_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Error Report", data=csv, file_name="validation_errors.csv", mime="text/csv"
                )
            else:
                st.success("Validation completed: no errors found!")
                if st.button("Import Data", type="primary"):
                    run_import(workbook)


if __name__ == "__main__":
    main()