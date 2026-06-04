# src/excel_reader.py
import openpyxl  # type: ignore
import os


def read_store_data(input_file_path, inspection_template_path):
    """
    1. Reads Store/Property data from the browsed Input file.
    2. Injects them into the Inspection Sheet template to let formulas compute.
    3. Extracts the computed tables and values.
    """
    # Grab the Store Code from the filename (e.g., Z943)
    filename = os.path.basename(input_file_path)
    store_code = filename.split('-')[0].strip()

    # --- STEP 1: READ FROM INPUT DATA EXCEL ---
    try:
        wb_input = openpyxl.load_workbook(input_file_path, data_only=True)

        # We look for the exact name or a lowercase/stripped version of it
        sheet_names_clean = {sheet.replace(
            " ", "").lower(): sheet for sheet in wb_input.sheetnames}

        # This matches 'property&storedetails', 'propertyandstoredetails', etc.
        target_variant_1 = "property&storedetails"
        target_variant_2 = "propertyandstoredetails"

        if target_variant_1 in sheet_names_clean:
            ws_details = wb_input[sheet_names_clean[target_variant_1]]
        elif target_variant_2 in sheet_names_clean:
            ws_details = wb_input[sheet_names_clean[target_variant_2]]
        else:
            # Fallback to whatever sheet is at index 0 if the name completely changes
            ws_details = wb_input.worksheets[0]

    except Exception as e:
        raise RuntimeError(f"Error reading Input Excel file: {e}")

    store_area = ws_details["E56"].value
    property_area = ws_details["E9"].value

    # Read the 10 mass values from E83:E92
    mass_values = []
    for row in ws_details.iter_rows(min_row=83, max_row=92, min_col=5, max_col=5, values_only=True):
        mass_values.append(row[0] if row[0] is not None else 0)

    wb_input.close()

    # --- STEP 2: WRITE INTO INSPECTION SHEET (FORMULA ENGINE) ---
    if not os.path.exists(inspection_template_path):
        raise FileNotFoundError(
            f"Inspection template sheet missing at: {inspection_template_path}")

    wb_inspect = openpyxl.load_workbook(
        inspection_template_path, data_only=False)  # Keep formulas alive

    # (i) Sheet 1: Occupancy load Calculation
    ws_occ = wb_inspect["Occupancy load Calculation"]
    for row_idx in range(2, 15):  # D2 to D14
        ws_occ[f"D{row_idx}"] = store_area

    # (ii) Sheet 2: Fire Load Calculation
    ws_fire = wb_inspect["Fire Load Calculation"]
    # Paste masses into E2:E11
    for i, mass in enumerate(mass_values):
        ws_fire[f"E{i+2}"] = mass
    # Enter area in F13
    ws_fire["F13"] = store_area

    # Save to a temporary workbook and re-open with data_only=True to let openpyxl read evaluated values
    temp_inspect_path = "data/temp_calculated_inspection.xlsx"
    wb_inspect.save(temp_inspect_path)
    wb_inspect.close()

    # --- STEP 3: READ EVALUATED RESULTS FROM THE TEMPORARY SHEET ---
    wb_evaluated = openpyxl.load_workbook(temp_inspect_path, data_only=True)

    extracted_data = {
        "store_code": store_code,
        "store_area": store_area,
        "property_area": property_area,
        "occupancy_table": [],
        "fire_load_density": None,
        "total_fire_load": None
    }

    # Extract Sheet 1 grid data (Area Occupied F2:F15, Occupancy Load I2:I15)
    ws_occ_eval = wb_evaluated["Occupancy load Calculation"]
    for r in range(2, 16):  # rows 2 to 15
        extracted_data["occupancy_table"].append({
            "area_occupied": ws_occ_eval[f"F{r}"].value,
            "occupancy_load": ws_occ_eval[f"I{r}"].value
        })

    # Extract Sheet 2 grid values (Fire Load Density in F14, Total Fire Load if needed)
    ws_fire_eval = wb_evaluated["Fire Load Calculation"]
    extracted_data["fire_load_density"] = ws_fire_eval["F14"].value

    wb_evaluated.close()

    # Clean up the temporary calculated file safely
    try:
        os.remove(temp_inspect_path)
    except:
        pass

    return extracted_data
