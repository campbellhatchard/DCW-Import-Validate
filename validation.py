"""
validation.py

This module implements the validation logic for the DCW Import & Validation Tool.  It
contains enumerations for allowed values, helper functions to perform cross‑tab
checks and a `validate_workbook` function which accepts a dictionary of
pandas DataFrames (keyed by sheet name) and returns a list of error
dictionaries.  Each error dictionary contains the sheet name, field name
and a descriptive message.

The validation rules follow the specification outlined in the accompanying
document (see report.md in the root of this project).  Mandatory fields,
uniqueness constraints, enumerated values and cross references are checked
for each tab in the expected order.
"""

from __future__ import annotations

import datetime
from typing import Dict, List, Tuple, Set

import pandas as pd

# Enumerations and lookup lists.  These come directly from the validation
# rules document.  They are defined here in uppercase lists for easy reuse.

ORG_UNIT_TYPES: Set[str] = {
    "Company",
    "Elimination Company",
    "Balance Sheet",
    "Business Unit",
    "Division",
    "Cost Center",
    "Department",
    "Plant",
    "Warehouse",
    "Region",
    "Inventory Warehouse",
}

ORG_UNIT_STRUCTURES: Set[str] = {
    "Inventory Org Unit",
    "Job Cost Org Unit",
    "Joint Venture",
    "Manufacturing Plant",
    "Purchasing Org Unit",
    "Sales Org Unit",
    "Shared Service Org Unit",
    "WMS Org Unit",
}

ZONE_TYPES: Set[str] = {
    "Inbound Receiving",
    "Inspection",
    "Location",
    "Outbound Packing",
    "Outbound Picking",
    "Outbound Shipping",
    "Replenishment",
}

LOCATION_TYPES: Set[str] = {
    "Receiving",
    "Fixed Picking",
    "Packing",
    "Shipping",
    "Replenishment",
    "Inspection",
}

ITEM_TYPES: Set[str] = {
    "Battery",
    "BFC Solutions",
    "Book",
    "CD",
    "Configured Materials",
    "Copy & Printer Paper",
    "Desk & Tables",
    "Digital Audio",
    "Digital Video",
    "DVD",
    "Electronics",
    "Fabricated Part",
    "Fee",
    "Finished Goods",
    "Food & Beverage",
    "Garden Ranch",
    "Magazine",
    "Maintenance & Repair",
    "MFG-Assembly",
    "MFG-Bundle",
    "MFG-Feature",
    "MFG-Finished Product",
    "MFG-Kit",
    "MFG-Part",
    "MFG-Phantom",
    "MFG-Raw Material",
    "Miscellaneous",
    "Office Accessories",
    "Office Chairs",
    "Office Supplies",
    "Operating Supplies",
    "Packaging",
    "Premier Barn Garage",
    "Premier Pro Garage",
    "Premier Pro Ranch",
    "Premier Tall Barn",
    "Production Tools & Fixtures",
    "Resale",
    "Service",
    "Wires & Cables",
}

# Valid UOM codes.  This list consolidates the values used across several tabs.
UOMS: Set[str] = {
    "ACRE", "ARES", "BA", "BG", "BO", "BOLT", "BX", "CA", "CASE", "CFT", "CIN",
    "CL", "CM", "CR", "CT", "CUP", "CYC", "DL", "DM", "DRUM", "EA", "EMP", "FOZ",
    "FT", "G", "GAL", "HA", "HL", "IN", "KG", "KM", "L", "LB", "M", "M2", "M3", "MG",
    "MI", "ML", "MM", "MM3", "OZ", "PALLET", "PF", "PK", "PT", "QT", "REAM", "RL",
    "ROLL", "SFT", "SH", "SHEET", "SIN", "SQF", "SYD", "TB", "TBSP", "TNE", "TON",
    "TSP", "YD",
}

UOM_GROUPS: Set[str] = {
    "Area (metric)", "Area (US)", "Length (metric)", "Length (US)", "Mass (Metric)",
    "Mass (US)", "Pressure", "Quantity", "Temperature", "Volume (metric)",
    "Volume - Liquid (US)", "Volume - Solid (US)",
}

CONVERSION_TYPES: Set[str] = {"Item Specific", "Standard", "Supplier-Item Specific"}

ITEM_CROSS_REF_UPC_TYPES: Set[str] = {
    "UPC or EAN Code (13 digit)",
    "Customer Item ID",
    "Manufacturer",
    "GTIN",
}

ITEM_CROSS_REF_SUPPLIER_TYPE: str = "Supplier Item ID"

MAKE_TYPES: Set[str] = {"Kit", "Manufactured"}
BUY_TYPES: Set[str] = {"Purchased", "Service"}
INVENTORY_VALUATION_METHODS: Set[str] = {"Standard", "Average", "Last In"}
GL_ITEM_CLASSES: Set[str] = {
    "Accesories",  # spelled as in the specification
    "Assembly",
    "Computer Accessories",
    "Copy & Printer Paper",
    "Desks & Tables",
    "Electronics",
    "Fabricated Part",
    "Finished Good",
    "Finished Product",
    "Food & Beverage",
    "Office Chairs",
    "Office Furniture",
    "Office Supplies",
    "Packaging Material",
    "Part",
    "Raw Materials",
    "Whiteboards",
}
GOODS_SERVICES: Set[str] = {"Services", "Goods"}
SERIAL_CAPTURE_METHODS: Set[str] = {"Location"}
SERIAL_CYCLE_COUNT_OVERRIDE_RULES: Set[str] = {"Always Capture", "Capture on Unexpected Quantity"}
TOLERANCE_TYPES: Set[str] = {"Count Percentage", "Count Quantity"}
REPLENISHMENT_METHODS: Set[str] = {"Maximum Quantity", "Replenishment Quantity"}


def _is_blank(value) -> bool:
    """Return True if a cell value is considered blank."""
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def validate_workbook(workbook: Dict[str, pd.DataFrame]) -> List[Dict[str, str]]:
    """
    Validate all sheets in the workbook.  Return a list of error
    dictionaries with keys: 'table', 'field' and 'message'.  The
    workbook should be read with header row already processed; row 0
    corresponds to the first data row (original row 2 in Excel).
    """
    errors: List[Dict[str, str]] = []

    # Normalise all dataframes: ensure string values and no NaN
    dataframes: Dict[str, pd.DataFrame] = {}
    for name, df in workbook.items():
        # Convert column names to strings
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        # Fill NaN with empty string for validation
        df = df.fillna("")
        dataframes[name] = df

    # Build context for cross references
    org_unit_numbers: Set[str] = set()
    warehouse_names: Set[str] = set()
    contact_types: Set[str] = set()
    # Additional sets populated later
    zone_names: Set[str] = set()
    location_pairs: Set[Tuple[str, str]] = set()
    location_groups: Set[Tuple[str, str]] = set()
    item_numbers: Set[str] = set()
    item_org_pairs: Set[Tuple[str, str]] = set()
    customer_names: Set[str] = set()
    supplier_names: Set[str] = set()
    job_names: Set[str] = set()

    # Helper for adding error
    def add_error(table: str, field: str, message: str) -> None:
        errors.append({"table": table, "field": field, "message": message})

    # --- Validate Org Units ---
    org_df = dataframes.get("Org Units")
    if org_df is None:
        add_error("Org Units", "sheet", "Org Units sheet is missing from workbook")
    else:
        if org_df.shape[0] == 0:
            add_error("Org Units", "", "At least one org unit must exist")
        # Keep track of unique keys
        seen_org_keys: Set[Tuple[str, str]] = set()
        for idx, row in org_df.iterrows():
            excel_row = idx + 2  # Excel row number (header row is row 1)
            ou_number = str(row.get("Org Unit Number", "")).strip()
            wh_name = str(row.get("Warehouse Name", "")).strip()
            contact_type = str(row.get("Contact Type", "")).strip()
            org_unit_type = str(row.get("Org Unit Type", "")).strip()
            org_unit_function = str(row.get("Org Unit Function", "")).strip()
            contact_type_group = str(row.get("Contact Type Group", "")).strip()
            # Mandatory fields
            if _is_blank(ou_number):
                add_error("Org Units", "Org Unit Number", f"Row {excel_row}: Org Unit Number is mandatory")
            if _is_blank(wh_name):
                add_error("Org Units", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory")
            if _is_blank(contact_type):
                add_error("Org Units", "Contact Type", f"Row {excel_row}: Contact Type is mandatory")
            # Unique key check
            key = (ou_number, wh_name)
            if ou_number and wh_name:
                if key in seen_org_keys:
                    add_error("Org Units", "Org Unit Number/Warehouse Name", f"Row {excel_row}: Duplicate Org Unit Number and Warehouse Name combination: {key}")
                else:
                    seen_org_keys.add(key)
                    org_unit_numbers.add(ou_number)
                    warehouse_names.add(wh_name)
            # Collect contact type
            if contact_type:
                contact_types.add(contact_type)
            # Enumerations
            if org_unit_type and org_unit_type not in ORG_UNIT_TYPES:
                add_error("Org Units", "Org Unit Type", f"Row {excel_row}: Invalid Org Unit Type '{org_unit_type}'")
            # Org Unit Structure may be stored in Org Unit Function column
            if org_unit_function and org_unit_function not in ORG_UNIT_STRUCTURES:
                add_error("Org Units", "Org Unit Function", f"Row {excel_row}: Invalid Org Unit Structure/Function '{org_unit_function}'")
            # Contact Type Group must equal IO
            if contact_type and contact_type_group and contact_type_group != "IO":
                add_error("Org Units", "Contact Type Group", f"Row {excel_row}: Contact Type Group must be 'IO' when Contact Type is provided")

    # --- Validate Org Unit Relationships ---
    rel_df = dataframes.get("Org Unit Relationships")
    if rel_df is not None and rel_df.shape[0] > 0:
        for idx, row in rel_df.iterrows():
            excel_row = idx + 2
            wh_number = str(row.get("Warehouse Number", "")).strip()
            parent_number = str(row.get("Parent OrgUnit Number", "")).strip()
            # Mandatory fields
            if _is_blank(wh_number):
                add_error("Org Unit Relationships", "Warehouse Number", f"Row {excel_row}: Warehouse Number is mandatory")
            if _is_blank(parent_number):
                add_error("Org Unit Relationships", "Parent OrgUnit Number", f"Row {excel_row}: Parent OrgUnit Number is mandatory")
            # Referential integrity
            if wh_number and wh_number not in org_unit_numbers:
                add_error("Org Unit Relationships", "Warehouse Number", f"Row {excel_row}: Warehouse Number '{wh_number}' does not exist in Org Units")
            if parent_number and parent_number not in org_unit_numbers:
                add_error("Org Unit Relationships", "Parent OrgUnit Number", f"Row {excel_row}: Parent OrgUnit Number '{parent_number}' does not exist in Org Units")

    # --- Validate Customers ---
    cust_df = dataframes.get("Customers")
    if cust_df is not None and cust_df.shape[0] > 0:
        for idx, row in cust_df.iterrows():
            excel_row = idx + 2
            # Determine if row has any data
            if all(_is_blank(v) for v in row.values):
                continue
            contact_type = str(row.get("Contact Type", "")).strip()
            name = str(row.get("Customer Name", "")).strip()
            address = str(row.get("Address", "")).strip()
            city = str(row.get("City", "")).strip()
            state = str(row.get("State", "")).strip()
            zip_code = str(row.get("Zip Code", "")).strip()
            # Mandatory fields
            if _is_blank(contact_type):
                add_error("Customers", "Contact Type", f"Row {excel_row}: Contact Type is mandatory")
            if _is_blank(name):
                add_error("Customers", "Customer Name", f"Row {excel_row}: Customer Name is mandatory")
            else:
                customer_names.add(name)
            # If customer created, address fields required
            if name:
                if _is_blank(address):
                    add_error("Customers", "Address", f"Row {excel_row}: Address is mandatory for Customer '{name}'")
                if _is_blank(city):
                    add_error("Customers", "City", f"Row {excel_row}: City is mandatory for Customer '{name}'")
                if _is_blank(state):
                    add_error("Customers", "State", f"Row {excel_row}: State is mandatory for Customer '{name}'")
                if _is_blank(zip_code):
                    add_error("Customers", "Zip Code", f"Row {excel_row}: Zip Code is mandatory for Customer '{name}'")

    # --- Validate Suppliers ---
    supp_df = dataframes.get("Suppliers")
    if supp_df is not None and supp_df.shape[0] > 0:
        for idx, row in supp_df.iterrows():
            excel_row = idx + 2
            # skip blank rows
            if all(_is_blank(v) for v in row.values):
                continue
            contact_type = str(row.get("Contact Type", "")).strip()
            name = str(row.get("Supplier Name", "")).strip()
            address = str(row.get("Address", "")).strip()
            city = str(row.get("City", "")).strip()
            state = str(row.get("State", "")).strip()
            zip_code = str(row.get("Zip Code", "")).strip()
            if _is_blank(contact_type):
                add_error("Suppliers", "Contact Type", f"Row {excel_row}: Contact Type is mandatory")
            if _is_blank(name):
                add_error("Suppliers", "Supplier Name", f"Row {excel_row}: Supplier Name is mandatory")
            else:
                supplier_names.add(name)
            # If supplier created, address fields required
            if name:
                if _is_blank(address):
                    add_error("Suppliers", "Address", f"Row {excel_row}: Address is mandatory for Supplier '{name}'")
                if _is_blank(city):
                    add_error("Suppliers", "City", f"Row {excel_row}: City is mandatory for Supplier '{name}'")
                if _is_blank(state):
                    add_error("Suppliers", "State", f"Row {excel_row}: State is mandatory for Supplier '{name}'")
                if _is_blank(zip_code):
                    add_error("Suppliers", "Zip Code", f"Row {excel_row}: Zip Code is mandatory for Supplier '{name}'")

    # --- Validate Zones ---
    zones_df = dataframes.get("Zones")
    if zones_df is None or zones_df.shape[0] == 0:
        add_error("Zones", "", "At least one zone must exist")
    else:
        seen_zone_names: Set[str] = set()
        for idx, row in zones_df.iterrows():
            excel_row = idx + 2
            wh = str(row.get("Warehouse Name", "")).strip()
            wh_filter = str(row.get("Warehouse Name Filter", "")).strip()
            zone_name = str(row.get("Zone Name", "")).strip()
            zone_type = str(row.get("Zone Type", "")).strip()
            # Mandatory fields
            if _is_blank(wh):
                add_error("Zones", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory")
            if _is_blank(zone_name):
                add_error("Zones", "Zone Name", f"Row {excel_row}: Zone Name is mandatory")
            if _is_blank(zone_type):
                add_error("Zones", "Zone Type", f"Row {excel_row}: Zone Type is mandatory")
            # Uniqueness
            if zone_name:
                if zone_name in seen_zone_names:
                    add_error("Zones", "Zone Name", f"Row {excel_row}: Duplicate Zone Name '{zone_name}'")
                else:
                    seen_zone_names.add(zone_name)
                    zone_names.add(zone_name)
            # Valid values
            if zone_type and zone_type not in ZONE_TYPES:
                add_error("Zones", "Zone Type", f"Row {excel_row}: Invalid Zone Type '{zone_type}'")
            # Cross references
            if wh and wh not in warehouse_names:
                add_error("Zones", "Warehouse Name", f"Row {excel_row}: Warehouse Name '{wh}' not found in Org Units")
            if wh_filter and wh_filter not in contact_types:
                add_error("Zones", "Warehouse Name Filter", f"Row {excel_row}: Warehouse Name Filter '{wh_filter}' is not a valid Contact Type from Org Units")

    # --- Validate ABC Codes ---
    abc_df = dataframes.get("ABC Codes")
    if abc_df is not None and abc_df.shape[0] > 0:
        for idx, row in abc_df.iterrows():
            excel_row = idx + 2
            code = str(row.get("ABC Code", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            wh_filter = str(row.get("Warehouse Name Filter", "")).strip()
            # Check code
            if code and code not in {"A", "B", "C"}:
                add_error("ABC Codes", "ABC Code", f"Row {excel_row}: Invalid ABC Code '{code}'")
            if _is_blank(code):
                add_error("ABC Codes", "ABC Code", f"Row {excel_row}: ABC Code is mandatory")
            # Cross references
            if wh and wh not in warehouse_names:
                add_error("ABC Codes", "Warehouse Name", f"Row {excel_row}: Warehouse Name '{wh}' not found in Org Units")
            if wh_filter and wh_filter not in contact_types:
                add_error("ABC Codes", "Warehouse Name Filter", f"Row {excel_row}: Warehouse Name Filter '{wh_filter}' is not a valid Contact Type from Org Units")

    # --- Validate Inventory Locations ---
    inv_loc_df = dataframes.get("Inventory Locations")
    if inv_loc_df is not None and inv_loc_df.shape[0] > 0:
        seen_loc_pairs: Set[Tuple[str, str]] = set()
        for idx, row in inv_loc_df.iterrows():
            excel_row = idx + 2
            wh = str(row.get("Warehouse Name", "")).strip()
            zone = str(row.get("Zone Name", "")).strip()
            loc_name = str(row.get("Location Name", "")).strip()
            loc_type = str(row.get("Location Type", "")).strip()
            # Mandatory
            if _is_blank(wh):
                add_error("Inventory Locations", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory")
            if _is_blank(zone):
                add_error("Inventory Locations", "Zone Name", f"Row {excel_row}: Zone Name is mandatory")
            if _is_blank(loc_name):
                add_error("Inventory Locations", "Location Name", f"Row {excel_row}: Location Name is mandatory")
            if _is_blank(loc_type):
                add_error("Inventory Locations", "Location Type", f"Row {excel_row}: Location Type is mandatory")
            # Unique combination
            key = (wh, loc_name)
            if wh and loc_name:
                if key in seen_loc_pairs:
                    add_error("Inventory Locations", "Location Name", f"Row {excel_row}: Duplicate Location Name '{loc_name}' in Warehouse '{wh}'")
                else:
                    seen_loc_pairs.add(key)
                    location_pairs.add(key)
            # Valid location type
            if loc_type and loc_type not in LOCATION_TYPES:
                add_error("Inventory Locations", "Location Type", f"Row {excel_row}: Invalid Location Type '{loc_type}'")
            # Cross references
            if wh and wh not in warehouse_names:
                add_error("Inventory Locations", "Warehouse Name", f"Row {excel_row}: Warehouse Name '{wh}' not found in Org Units")
            if zone and zone not in zone_names:
                add_error("Inventory Locations", "Zone Name", f"Row {excel_row}: Zone Name '{zone}' not found in Zones")

    # --- Validate Location Groups ---
    loc_groups_df = dataframes.get("Location Groups")
    if loc_groups_df is not None and loc_groups_df.shape[0] > 0:
        for idx, row in loc_groups_df.iterrows():
            excel_row = idx + 2
            wh = str(row.get("Warehouse Name", "")).strip()
            loc_group = str(row.get("Location Group", "")).strip()
            loc_name = str(row.get("Location Name", "")).strip()
            # Skip blank rows
            if all(_is_blank(v) for v in row.values):
                continue
            # Mandatory fields when creating a record
            if _is_blank(wh):
                add_error("Location Groups", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory")
            if _is_blank(loc_group):
                add_error("Location Groups", "Location Group", f"Row {excel_row}: Location Group is mandatory")
            if _is_blank(loc_name):
                add_error("Location Groups", "Location Name", f"Row {excel_row}: Location Name is mandatory")
            # Cross references
            if wh and wh not in warehouse_names:
                add_error("Location Groups", "Warehouse Name", f"Row {excel_row}: Warehouse Name '{wh}' not found in Org Units")
            if loc_name and (wh, loc_name) not in location_pairs:
                add_error("Location Groups", "Location Name", f"Row {excel_row}: Location Name '{loc_name}' for Warehouse '{wh}' not found in Inventory Locations")
            # Add to location_groups set for later cross reference
            if wh and loc_group:
                location_groups.add((wh, loc_group))

    # --- Validate Item Global Master ---
    item_master_df = dataframes.get("Item Global Master")
    if item_master_df is not None and item_master_df.shape[0] > 0:
        seen_items: Set[str] = set()
        for idx, row in item_master_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            desc = str(row.get("Item Description", "")).strip()
            item_type = str(row.get("Item Type", "")).strip()
            uom = str(row.get("Primary UOM", "")).strip()
            effective_from = row.get("Effective From")
            # Unique item
            if _is_blank(item_num):
                add_error("Item Global Master", "Item Number", f"Row {excel_row}: Item Number is mandatory")
            else:
                if item_num in seen_items:
                    add_error("Item Global Master", "Item Number", f"Row {excel_row}: Duplicate Item Number '{item_num}'")
                else:
                    seen_items.add(item_num)
                    item_numbers.add(item_num)
            # Mandatory fields
            if _is_blank(desc):
                add_error("Item Global Master", "Item Description", f"Row {excel_row}: Item Description is mandatory for Item '{item_num}'")
            if _is_blank(item_type):
                add_error("Item Global Master", "Item Type", f"Row {excel_row}: Item Type is mandatory for Item '{item_num}'")
            elif item_type not in ITEM_TYPES:
                add_error("Item Global Master", "Item Type", f"Row {excel_row}: Invalid Item Type '{item_type}' for Item '{item_num}'")
            if _is_blank(uom):
                add_error("Item Global Master", "Primary UOM", f"Row {excel_row}: Primary UOM is mandatory for Item '{item_num}'")
            elif uom not in UOMS:
                add_error("Item Global Master", "Primary UOM", f"Row {excel_row}: Invalid Primary UOM '{uom}' for Item '{item_num}'")
            # If Effective From is blank, default to current date (no error)
            # (we leave defaulting to import logic; not stored here)

    # --- Validate Item Org Units ---
    item_org_df = dataframes.get("Item Org Units")
    if item_org_df is not None and item_org_df.shape[0] > 0:
        seen_item_org: Set[Tuple[str, str]] = set()
        for idx, row in item_org_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            loc_group = str(row.get("Location Group", "")).strip()
            primary_loc = str(row.get("Primary Location Name", "")).strip()
            zone = str(row.get("Zone Name", "")).strip()
            abc = str(row.get("ABC Code", "")).strip()
            make_type = str(row.get("Make Type", "")).strip()
            buy_type = str(row.get("Buy Type", "")).strip()
            inv_val = str(row.get("Inventory Valuation Method", "")).strip()
            gl_class = str(row.get("GL Item Class", "")).strip()
            goods_services = str(row.get("Goods or Services Indicator", "")).strip()
            serial_capture = str(row.get("Serial Capture Method", "")).strip()
            serial_override = str(row.get("Serial Cycle Count Override Rule", "")).strip()
            # Unique combination
            if not _is_blank(item_num) and not _is_blank(wh):
                key = (item_num, wh)
                if key in seen_item_org:
                    add_error("Item Org Units", "Item Number/Warehouse Name", f"Row {excel_row}: Duplicate Item Number '{item_num}' in Warehouse '{wh}'")
                else:
                    seen_item_org.add(key)
                    item_org_pairs.add(key)
            # Cross references
            if item_num and item_num not in item_numbers:
                add_error("Item Org Units", "Item Number", f"Row {excel_row}: Item Number '{item_num}' does not exist in Item Global Master")
            if wh and wh not in warehouse_names:
                add_error("Item Org Units", "Warehouse Name", f"Row {excel_row}: Warehouse Name '{wh}' does not exist in Org Units")
            # Optional enumerations
            if make_type and make_type not in MAKE_TYPES:
                add_error("Item Org Units", "Make Type", f"Row {excel_row}: Invalid Make Type '{make_type}' for Item '{item_num}'")
            if buy_type and buy_type not in BUY_TYPES:
                add_error("Item Org Units", "Buy Type", f"Row {excel_row}: Invalid Buy Type '{buy_type}' for Item '{item_num}'")
            if inv_val and inv_val not in INVENTORY_VALUATION_METHODS:
                add_error("Item Org Units", "Inventory Valuation Method", f"Row {excel_row}: Invalid Inventory Valuation Method '{inv_val}' for Item '{item_num}'")
            if gl_class and gl_class not in GL_ITEM_CLASSES:
                add_error("Item Org Units", "GL Item Class", f"Row {excel_row}: Invalid GL Item Class '{gl_class}' for Item '{item_num}'")
            if goods_services and goods_services not in GOODS_SERVICES:
                add_error("Item Org Units", "Goods or Services Indicator", f"Row {excel_row}: Invalid Goods/Services Indicator '{goods_services}' for Item '{item_num}'")
            if serial_capture and serial_capture not in SERIAL_CAPTURE_METHODS:
                add_error("Item Org Units", "Serial Capture Method", f"Row {excel_row}: Invalid Serial Capture Method '{serial_capture}' for Item '{item_num}'")
            if serial_override and serial_override not in SERIAL_CYCLE_COUNT_OVERRIDE_RULES:
                add_error("Item Org Units", "Serial Cycle Count Override Rule", f"Row {excel_row}: Invalid Serial Cycle Count Override Rule '{serial_override}' for Item '{item_num}'")
            if abc and abc not in {"A", "B", "C"}:
                add_error("Item Org Units", "ABC Code", f"Row {excel_row}: Invalid ABC Code '{abc}' for Item '{item_num}'")
            # Cross references to other tables
            if primary_loc:
                if (wh, primary_loc) not in location_pairs:
                    add_error("Item Org Units", "Primary Location Name", f"Row {excel_row}: Primary Location '{primary_loc}' does not exist for Warehouse '{wh}'")
            if loc_group:
                if (wh, loc_group) not in location_groups:
                    add_error("Item Org Units", "Location Group", f"Row {excel_row}: Location Group '{loc_group}' does not exist for Warehouse '{wh}'")
            if zone:
                if zone not in zone_names:
                    add_error("Item Org Units", "Zone Name", f"Row {excel_row}: Zone Name '{zone}' does not exist in Zones")

    # --- Validate UOM Conversions ---
    uom_conv_df = dataframes.get("UOM Conversions")
    if uom_conv_df is not None and uom_conv_df.shape[0] > 0:
        for idx, row in uom_conv_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            from_qty = row.get("Default From UOM Qty")
            from_uom = str(row.get("From UOM", "")).strip()
            to_uom = str(row.get("To UOM", "")).strip()
            conv_type = str(row.get("Conversion Type", "")).strip()
            uom_group = str(row.get("Unit Of Measure Group", "")).strip()
            dim_uom = str(row.get("Dimension UOM", "")).strip()
            weight_uom = str(row.get("Weight UOM", "")).strip()
            # If item number provided, must exist
            if item_num and item_num not in item_numbers:
                add_error("UOM Conversions", "Item Number", f"Row {excel_row}: Item Number '{item_num}' does not exist in Item Global Master")
            # From quantity must equal 1
            if from_qty != "" and from_qty != 1 and from_qty != 1.0:
                add_error("UOM Conversions", "Default From UOM Qty", f"Row {excel_row}: Default From UOM Qty must equal 1")
            # Conversion type
            if conv_type and conv_type not in CONVERSION_TYPES:
                add_error("UOM Conversions", "Conversion Type", f"Row {excel_row}: Invalid Conversion Type '{conv_type}'")
            # UOM group
            if uom_group and uom_group not in UOM_GROUPS:
                add_error("UOM Conversions", "Unit Of Measure Group", f"Row {excel_row}: Invalid UOM Group '{uom_group}'")
            # UOM values
            for field_name, val in [("From UOM", from_uom), ("To UOM", to_uom), ("Dimension UOM", dim_uom), ("Weight UOM", weight_uom)]:
                if val and val not in UOMS:
                    add_error("UOM Conversions", field_name, f"Row {excel_row}: Invalid UOM '{val}'")

    # --- Validate Item Capacities ---
    cap_df = dataframes.get("Item Capacities")
    if cap_df is not None and cap_df.shape[0] > 0:
        for idx, row in cap_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            loc_type = str(row.get("Location Type", "")).strip()
            cap_qty = str(row.get("Capacity Quantity", "")).strip()
            cap_uom = str(row.get("Capacity UOM", "")).strip()
            if not _is_blank(item_num):
                # mandatory additional fields
                if _is_blank(wh):
                    add_error("Item Capacities", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory when Item Number is provided")
                if _is_blank(loc_type):
                    add_error("Item Capacities", "Location Type", f"Row {excel_row}: Location Type is mandatory when Item Number is provided")
                if _is_blank(cap_qty):
                    add_error("Item Capacities", "Capacity Quantity", f"Row {excel_row}: Capacity Quantity is mandatory when Item Number is provided")
                if _is_blank(cap_uom):
                    add_error("Item Capacities", "Capacity UOM", f"Row {excel_row}: Capacity UOM is mandatory when Item Number is provided")
                # Cross references
                if item_num and wh and (item_num, wh) not in item_org_pairs:
                    add_error("Item Capacities", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item Number '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
                # Location type allowed
                if loc_type and loc_type not in LOCATION_TYPES:
                    add_error("Item Capacities", "Location Type", f"Row {excel_row}: Invalid Location Type '{loc_type}'")
                # UOM allowed
                if cap_uom and cap_uom not in UOMS:
                    add_error("Item Capacities", "Capacity UOM", f"Row {excel_row}: Invalid Capacity UOM '{cap_uom}'")

    # --- Validate Item Tolerances ---
    tol_df = dataframes.get("Item Tolerances")
    if tol_df is not None and tol_df.shape[0] > 0:
        for idx, row in tol_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            tol_type = str(row.get("Tolerance Type", "")).strip()
            tol_upper = str(row.get("Tolerance Value Upper", "")).strip()
            tol_lower = str(row.get("Tolerance Value Lower", "")).strip()
            if not _is_blank(item_num):
                # mandatory
                if _is_blank(wh):
                    add_error("Item Tolerances", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory when Item Number is provided")
                if _is_blank(tol_type):
                    add_error("Item Tolerances", "Tolerance Type", f"Row {excel_row}: Tolerance Type is mandatory when Item Number is provided")
                if _is_blank(tol_upper):
                    add_error("Item Tolerances", "Tolerance Value Upper", f"Row {excel_row}: Tolerance Value Upper is mandatory when Item Number is provided")
                if _is_blank(tol_lower):
                    add_error("Item Tolerances", "Tolerance Value Lower", f"Row {excel_row}: Tolerance Value Lower is mandatory when Item Number is provided")
                # Cross references
                if item_num and wh and (item_num, wh) not in item_org_pairs:
                    add_error("Item Tolerances", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item Number '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
                # Tolerance type allowed
                if tol_type and tol_type not in TOLERANCE_TYPES:
                    add_error("Item Tolerances", "Tolerance Type", f"Row {excel_row}: Invalid Tolerance Type '{tol_type}'")

    # --- Validate Item Cross Ref UPC ---
    upc_df = dataframes.get("Item Cross Ref UPC")
    if upc_df is not None and upc_df.shape[0] > 0:
        for idx, row in upc_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()  # note: sheet has both filter and warehouse
            xref_type = str(row.get("Item Cross Reference Type", "")).strip()
            uom = str(row.get("Unit Of Measure", "")).strip()
            # Mandatory if item number present
            if not _is_blank(item_num):
                if _is_blank(wh):
                    add_error("Item Cross Ref UPC", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory for Item '{item_num}'")
                if _is_blank(xref_type):
                    add_error("Item Cross Ref UPC", "Item Cross Reference Type", f"Row {excel_row}: Item Cross Reference Type is mandatory for Item '{item_num}'")
                # Cross references
                if item_num and wh and (item_num, wh) not in item_org_pairs:
                    add_error("Item Cross Ref UPC", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
                # Allowed types (not Supplier Item ID)
                if xref_type and xref_type not in ITEM_CROSS_REF_UPC_TYPES:
                    add_error("Item Cross Ref UPC", "Item Cross Reference Type", f"Row {excel_row}: Invalid Item Cross Reference Type '{xref_type}' for Item '{item_num}'")
                if uom and uom not in UOMS:
                    add_error("Item Cross Ref UPC", "Unit Of Measure", f"Row {excel_row}: Invalid Unit Of Measure '{uom}'")

    # --- Validate Item Cross Ref Supplier ---
    supplier_ref_df = dataframes.get("Item Cross Ref Supplier")
    if supplier_ref_df is not None and supplier_ref_df.shape[0] > 0:
        for idx, row in supplier_ref_df.iterrows():
            excel_row = idx + 2
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            xref_type = str(row.get("Item Cross Reference Type", "")).strip()
            uom = str(row.get("Unit Of Measure", "")).strip()
            supplier_name = str(row.get("Supplier Name", "")).strip()
            # If item provided
            if not _is_blank(item_num):
                if _is_blank(wh):
                    add_error("Item Cross Ref Supplier", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory for Item '{item_num}'")
                if _is_blank(xref_type):
                    add_error("Item Cross Ref Supplier", "Item Cross Reference Type", f"Row {excel_row}: Item Cross Reference Type is mandatory for Item '{item_num}'")
                if _is_blank(supplier_name):
                    add_error("Item Cross Ref Supplier", "Supplier Name", f"Row {excel_row}: Supplier Name is mandatory for Item '{item_num}'")
                # Cross references
                if item_num and wh and (item_num, wh) not in item_org_pairs:
                    add_error("Item Cross Ref Supplier", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
                if supplier_name and supplier_name not in supplier_names:
                    add_error("Item Cross Ref Supplier", "Supplier Name", f"Row {excel_row}: Supplier Name '{supplier_name}' does not exist in Suppliers")
                # Allowed type: must be Supplier Item ID
                if xref_type and xref_type != ITEM_CROSS_REF_SUPPLIER_TYPE:
                    add_error("Item Cross Ref Supplier", "Item Cross Reference Type", f"Row {excel_row}: Item Cross Reference Type must be '{ITEM_CROSS_REF_SUPPLIER_TYPE}'")
                if uom and uom not in UOMS:
                    add_error("Item Cross Ref Supplier", "Unit Of Measure", f"Row {excel_row}: Invalid Unit Of Measure '{uom}'")

    # --- Validate Fixed Pick Items ---
    fixed_df = dataframes.get("Fixed Pick Items")
    if fixed_df is not None and fixed_df.shape[0] > 0:
        for idx, row in fixed_df.iterrows():
            excel_row = idx + 2
            wh = str(row.get("Warehouse Name", "")).strip()
            item_num = str(row.get("Item Number", "")).strip()
            loc_name = str(row.get("Location Name", "")).strip()
            owner_name = str(row.get("Owner Name", "")).strip()
            replen_method = str(row.get("Replenishment Method", "")).strip()
            repl_qty_uom = str(row.get("Replenishment Quantity UOM", "")).strip()
            min_onhand_uom = str(row.get("Minimum On Hand UOM", "")).strip()
            max_stock_uom = str(row.get("Maximum Stock Level UOM", "")).strip()
            # If item populated, mandatory fields
            if not _is_blank(item_num):
                if _is_blank(wh):
                    add_error("Fixed Pick Items", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory when Item Number is provided")
                # Item Cross Reference Type column not present; ignoring
                # Cross references
                if wh and item_num and (item_num, wh) not in item_org_pairs:
                    add_error("Fixed Pick Items", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
                if loc_name and (wh, loc_name) not in location_pairs:
                    add_error("Fixed Pick Items", "Location Name", f"Row {excel_row}: Location Name '{loc_name}' for Warehouse '{wh}' not found in Inventory Locations")
                if replen_method and replen_method not in REPLENISHMENT_METHODS:
                    add_error("Fixed Pick Items", "Replenishment Method", f"Row {excel_row}: Invalid Replenishment Method '{replen_method}'")
                # Owner name cross reference
                if owner_name and owner_name not in customer_names:
                    add_error("Fixed Pick Items", "Owner Name", f"Row {excel_row}: Owner Name '{owner_name}' does not exist in Customers")
                # UOM validations
                for field_name, val in [
                    ("Replenishment Quantity UOM", repl_qty_uom),
                    ("Minimum On Hand UOM", min_onhand_uom),
                    ("Maximum Stock Level UOM", max_stock_uom),
                ]:
                    if val and val not in UOMS:
                        add_error("Fixed Pick Items", field_name, f"Row {excel_row}: Invalid UOM '{val}'")

    # --- Validate Jobs ---
    jobs_df = dataframes.get("Jobs")
    if jobs_df is not None and jobs_df.shape[0] > 0:
        seen_jobs: Set[str] = set()
        for idx, row in jobs_df.iterrows():
            excel_row = idx + 2
            name = str(row.get("Job Name", "")).strip()
            # Determine if row has any data
            if all(_is_blank(v) for v in row.values):
                continue
            if _is_blank(name):
                add_error("Jobs", "Job Name", f"Row {excel_row}: Job Name is mandatory")
            else:
                if name in seen_jobs:
                    add_error("Jobs", "Job Name", f"Row {excel_row}: Duplicate Job Name '{name}'")
                else:
                    seen_jobs.add(name)
                    job_names.add(name)

    # --- Validate Serials ---
    serials_df = dataframes.get("Serials")
    if serials_df is not None and serials_df.shape[0] > 1:  # row 1 is sample data
        # Drop first row (sample) if it contains sample data: treat always dropping index 0
        df_serials = serials_df.iloc[1:].reset_index(drop=True)
        seen_serial_item: Set[Tuple[str, str]] = set()
        for idx, row in df_serials.iterrows():
            excel_row = idx + 3  # plus 3 because we skipped one data row and header row
            serial_num = str(row.get("Serial Number", "")).strip()
            item_num = str(row.get("Item Number", "")).strip()
            wh = str(row.get("Warehouse Name", "")).strip()
            loc_name = str(row.get("Location Name", "")).strip()
            # Unique combination
            if serial_num and item_num:
                key = (serial_num, item_num)
                if key in seen_serial_item:
                    add_error("Serials", "Serial Number/Item Number", f"Row {excel_row}: Duplicate Serial Number '{serial_num}' for Item '{item_num}'")
                else:
                    seen_serial_item.add(key)
            # Cross references
            if item_num and wh and (item_num, wh) not in item_org_pairs:
                add_error("Serials", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
            if wh and loc_name and (wh, loc_name) not in location_pairs:
                add_error("Serials", "Warehouse Name/Location Name", f"Row {excel_row}: Warehouse '{wh}' and Location '{loc_name}' combination not found in Inventory Locations")

    # --- Validate Inventory Balances ---
    inv_bal_df = dataframes.get("Inventory Balances")
    if inv_bal_df is not None and inv_bal_df.shape[0] > 0:
        for idx, row in inv_bal_df.iterrows():
            excel_row = idx + 2
            wh = str(row.get("Warehouse Name", "")).strip()
            item_num = str(row.get("Item Number", "")).strip()
            loc_name = str(row.get("Location Name", "")).strip()
            qty_on_hand = row.get("Quantity On Hand")
            owner_name = str(row.get("Owner Name", "")).strip()
            job_name = str(row.get("Job Name", "")).strip()
            # Determine if row is blank
            if all(_is_blank(v) for v in row.values):
                continue
            # Mandatory fields
            if _is_blank(wh):
                add_error("Inventory Balances", "Warehouse Name", f"Row {excel_row}: Warehouse Name is mandatory")
            if _is_blank(item_num):
                add_error("Inventory Balances", "Item Number", f"Row {excel_row}: Item Number is mandatory")
            if _is_blank(loc_name):
                add_error("Inventory Balances", "Location Name", f"Row {excel_row}: Location Name is mandatory")
            if qty_on_hand == "" or qty_on_hand is None:
                add_error("Inventory Balances", "Quantity On Hand", f"Row {excel_row}: Quantity On Hand is mandatory")
            else:
                # Must be numeric and >=0
                try:
                    qty_val = float(qty_on_hand)
                    if qty_val < 0:
                        add_error("Inventory Balances", "Quantity On Hand", f"Row {excel_row}: Quantity On Hand must be greater than or equal to 0")
                except ValueError:
                    add_error("Inventory Balances", "Quantity On Hand", f"Row {excel_row}: Quantity On Hand '{qty_on_hand}' is not numeric")
            # Cross references
            if item_num and wh and (item_num, wh) not in item_org_pairs:
                add_error("Inventory Balances", "Item Number/Warehouse Name", f"Row {excel_row}: Combination of Item '{item_num}' and Warehouse '{wh}' not found in Item Org Units")
            if wh and loc_name and (wh, loc_name) not in location_pairs:
                add_error("Inventory Balances", "Warehouse Name/Location Name", f"Row {excel_row}: Warehouse '{wh}' and Location '{loc_name}' combination not found in Inventory Locations")
            if owner_name and owner_name not in customer_names:
                add_error("Inventory Balances", "Owner Name", f"Row {excel_row}: Owner Name '{owner_name}' does not exist in Customers")
            if job_name and job_name not in job_names:
                add_error("Inventory Balances", "Job Name", f"Row {excel_row}: Job Name '{job_name}' does not exist in Jobs")

    # --- Validate Users ---
    users_df = dataframes.get("Users")
    if users_df is not None and users_df.shape[0] > 1:
        # Drop first row (sample data) if there is sample
        df_users = users_df.iloc[1:].reset_index(drop=True)
        for idx, row in df_users.iterrows():
            excel_row = idx + 3  # header + sample row
            # Determine if row has any data
            if all(_is_blank(v) for v in row.values):
                continue
            login_id = str(row.get("LoginId", "")).strip()
            email = str(row.get("Email", "")).strip()
            first_name = str(row.get("FirstName", "")).strip()
            last_name = str(row.get("LastName", "")).strip()
            default_env = str(row.get("DefaultZone", "")).strip()
            default_lifecycle = str(row.get("DefaultLifecycle", "")).strip()
            # Mandatory fields
            if _is_blank(email):
                add_error("Users", "Email", f"Row {excel_row}: Email is mandatory")
            if _is_blank(login_id):
                add_error("Users", "LoginId", f"Row {excel_row}: LoginId/User Name is mandatory")
            if _is_blank(first_name):
                add_error("Users", "FirstName", f"Row {excel_row}: First Name is mandatory")
            if _is_blank(last_name):
                add_error("Users", "LastName", f"Row {excel_row}: Last Name is mandatory")
            if _is_blank(default_env):
                # If default environment is missing, will be defaulted; but report info
                add_error("Users", "DefaultZone", f"Row {excel_row}: Default Environment is missing and will default to current environment")
            if _is_blank(default_lifecycle):
                add_error("Users", "DefaultLifecycle", f"Row {excel_row}: Default Lifecycle is missing and will default to current environment lifecycle")

    return errors