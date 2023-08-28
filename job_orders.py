import openpyxl
import os
import datetime
from copy import copy

def get_top_left_cell_of_merged_region(worksheet, cell_address):
    """Identify the top-left cell of a merged region."""
    for merged_range in worksheet.merged_cells.ranges:
        if cell_address in merged_range:
            return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
    return worksheet[cell_address]

# Ensure the "New Job Orders" directory exists
output_dir = "New Job Orders"
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

# Load the main data workbook
source_wb = openpyxl.load_workbook("EVIDENTA COMANDA ALVEOPLAST.xlsx")
source_ws = source_wb["COMENZI ALVEOPLAST"]

start_row = int(input("Enter the starting row: "))
col_to_check = 'H'
for row in range(start_row, source_ws.max_row + 1):
    if not source_ws[col_to_check + str(row)].value:
        end_row = row - 1
        break

def copy_range(src_ws, dest_ws, src_range, dest_cell):
    rows = src_ws[src_range]
    dest_cell = dest_ws[dest_cell]
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            dest_cell.offset(i, j).value = cell.value
            if cell.has_style:
                dest_cell.offset(i, j)._style = copy(cell._style)

# Mapping for Job B
mapping = {
    "H": "N5",
    "I": "N6",
    "B": "N7",
    "P": ["N8", "N22"],
    "X": "N9",
    "T": "N13",
    "U": "N14",
    "W": "L16",
    "V": "G18",
    "S": "K19",
    "M": "N23",
    "O": "N24",  # Note: The logic for "N24" involves data from columns P and M
    "R": "G30"
}

# Mapping for Job A
mapping_job_a = {
    "H": "G5",
    "I": "G6",
    "B": "G7",
    "P": ["G8", "G22"],
    "X": "G9",
    "T": "G13",
    "U": "G14",
    "W": "L16",
    "V": "G18",
    "S": "G19",
    "M": "G23",
    "O": "G24",  # Note: The logic for "G24" involves data from columns P and M
    "R": "G30"
}

def sanitize_filename(filename):
    """Sanitize filenames."""
    sanitized_name = ''.join(char if char.isalnum() or char in [' ', '_'] else '-' for char in filename)
    return sanitized_name

def populate_and_save_template(job_a_row, job_b_row, num_pallets_a, remaining_sheets_a, num_pallets_b=None, remaining_sheets_b=None):
    template_wb = openpyxl.load_workbook('JO&EP Template.xlsx')
    template_ws = template_wb["SABLON"]

    for col, dest_cells in mapping_job_a.items():
        if not isinstance(dest_cells, list):
            dest_cells = [dest_cells]
        for dest_cell in dest_cells:
            top_left_cell = get_top_left_cell_of_merged_region(template_ws, dest_cell)
            # Logic for G9
            if dest_cell == "G9":
                calculated_value = 0  # or any default value you see fit
                try:
                    source_value_p = float(source_ws["P" + str(job_a_row)].value or 0)
                    source_value_x = float(source_ws["W" + str(job_a_row)].value or 0)
                    source_value_t = float(source_ws["T" + str(job_a_row)].value or 0)
                    source_value_u = float(source_ws["U" + str(job_a_row)].value or 0)
                    calculated_value = source_value_p * source_value_x / 1000 * source_value_t * source_value_u / 1000000
                    top_left_cell.value = calculated_value
                except ValueError:
                    top_left_cell.value = "Error"  # or any other default value or action you want to happen in case of an error
                top_left_cell.value = calculated_value

            # Logic for G24
            elif dest_cell == "G24":
                source_value_p = source_ws["P" + str(job_a_row)].value
                source_value_m = source_ws["M" + str(job_a_row)].value
                calculated_value = source_value_p / source_value_m if source_value_m else 0
                top_left_cell.value = calculated_value

            # Default logic
            else:
                source_value = source_ws[col + str(job_a_row)].value
                top_left_cell.value = source_value
    if job_b_row:  # Only proceed if there's a B job
        for col, dest_cells in mapping.items():
            if not isinstance(dest_cells, list):
                dest_cells = [dest_cells]
            for dest_cell in dest_cells:
                top_left_cell = get_top_left_cell_of_merged_region(template_ws, dest_cell)
                # Logic for N9
                if dest_cell == "N9":
                    calculated_value = 0  # or any default value you see fit
                    try:
                        source_value_p = float(source_ws["P" + str(job_b_row)].value or 0)
                        source_value_x = float(source_ws["W" + str(job_b_row)].value or 0)
                        source_value_t = float(source_ws["T" + str(job_b_row)].value or 0)
                        source_value_u = float(source_ws["U" + str(job_b_row)].value or 0)
                        calculated_value = source_value_p * source_value_x / 1000 * source_value_t * source_value_u / 1000000
                        top_left_cell.value = calculated_value
                    except ValueError:
                        top_left_cell.value = "Error"  # or any other default value or action you want to happen in case of an error
                    top_left_cell.value = calculated_value

                # Logic for N24
                elif dest_cell == "N24":
                    source_value_p = source_ws["P" + str(job_b_row)].value
                    source_value_m = source_ws["M" + str(job_b_row)].value
                    calculated_value = source_value_p / source_value_m if source_value_m else 0
                    top_left_cell.value = calculated_value

                # Default logic
                else:
                    source_value = source_ws[col + str(job_b_row)].value
                    top_left_cell.value = source_value
    # Mapping for values in column R
    r_mapping = {
        "A": {
            "C32": "100%",
            "D32": "Virgin PPC3600"
        },
        "B": {
            "C32": "78%",
            "C33": "15%",
            "C34": "2%",
            "C35": "5%",
            "D32": "Virgin PPC3600",
            "D33": "Carbonat",
            "D34": "Virgin PPC3600",
            "D35": "TALC"
        },
        "ESDt": {
            "C32": "27%",
            "C33": "20%",
            "C36": "53%",
            "D32": "Virgin PPC3600",
            "D33": "REGRANULAT",
            "D36": "PREMIX"
        },
        "D": {
            "C32": "43%",
            "C33": "55%",
            "C34": "2%",
            "D32": "Virgin PPC3600",
            "D33": "REGRANULAT",
            "D34": "CULOARE/" + source_ws["S" + str(job_a_row)].value
        },
        "E": {
            "C32": "13%",
            "C33": "45%",
            "C34": "2%",
            "C35": "20%",
            "C36": "20%",
            "D32": "Virgin PPC3600",
            "D33": "REGRANULAT A",
            "D34": "CULOARE/" + source_ws["S" + str(job_a_row)].value,
            "D35": "1 parte TALC + 3 parti Carbonat",
            "D36": "REGRANULAT B"
        }
    }

    # Check value in column R for job A
    r_value_a = source_ws["R" + str(job_a_row)].value
    if r_value_a in r_mapping:
        for cell, value in r_mapping[r_value_a].items():
            template_ws[cell] = value

    # If there's a B job, also check its R value
    if job_b_row:
        r_value_b = source_ws["R" + str(job_b_row)].value
        if r_value_b in r_mapping:
            for cell, value in r_mapping[r_value_b].items():
                template_ws[cell] = value

    total_sheets_a = int(source_ws.cell(row=job_a_row, column=16).value)
    total_sheets_b = int(source_ws.cell(row=job_b_row, column=16).value) if job_b_row else 0

    # Copy and paste the Job A template the correct number of times
    if job_a_row:
        for i in range(num_pallets_a):
            # Copy the Job A template
            src_range = "U16:AE49"
            dest_cell = "U" + str(51 + i * (49 - 16 + 2))
            copy_range(template_ws, template_ws, src_range, dest_cell)

    # Write the correct information into the Job A labels
    if job_a_row:
        for i in range(num_pallets_a):
            row_v72 = 72 + i * (49 - 16 + 2)
            row_v79 = 79 + i * (49 - 16 + 2)
            row_AA79 = 79 + i * (49 - 16 + 2)

            template_ws.cell(row=row_v72, column=22).value = source_ws.cell(row=current_a_row, column=7).value
            template_ws.cell(row=row_v79, column=22).value = i + 1
            template_ws.cell(row=row_AA79, column=27).value = num_pallets_a

            sheets_per_pallet = int(template_ws['G23'].value)
        
            if i == num_pallets_a - 2:
                remaining_sheets = total_sheets_a % sheets_per_pallet
                if remaining_sheets > 0:
                    template_ws.cell(row=row_v72, column=22).value = sheets_per_pallet + remaining_sheets
            else:
                template_ws.cell(row=row_v72, column=22).value = sheets_per_pallet

    # Copy and paste the Job B template the correct number of times
    if job_b_row and current_b_row:
        for i in range(num_pallets_b):
            # Copy the Job B template
            src_range = "AG16:AQ49"
            dest_cell = "AG" + str(51 + i * (49 - 16 + 2))
            copy_range(template_ws, template_ws, src_range, dest_cell)

    # Write the correct information into the Job B labels
    if job_b_row and current_b_row:
        for i in range(num_pallets_b):
            row_ah72 = 72 + i * (49 - 16 + 2)
            row_ah79 = 79 + i * (49 - 16 + 2)
            row_AL79 = 79 + i * (49 - 16 + 2)

            template_ws.cell(row=row_ah72, column=34 - 7).value = source_ws.cell(row=current_b_row, column=14).value
            template_ws.cell(row=row_ah79, column=34).value = i + 1
            template_ws.cell(row=row_AL79, column=39).value = num_pallets_b

            sheets_per_pallet = int(source_ws.cell(row=current_b_row, column=14).value)
        
            if i == num_pallets_b - 2:
                remaining_sheets = total_sheets_b % sheets_per_pallet
                if remaining_sheets > 0:
                    template_ws.cell(row=row_ah72, column=34).value = sheets_per_pallet + remaining_sheets
            else:
                template_ws.cell(row=row_ah72, column=34).value = sheets_per_pallet

    # Code for naming and saving the file
    client = source_ws["H" + str(job_a_row)].value or "Unknown"
    length = source_ws["S" + str(job_a_row)].value or "Unknown"
    width = source_ws["T" + str(job_a_row)].value or "Unknown"
    density = source_ws["V" + str(job_a_row)].value or "Unknown"
    thickness = source_ws["U" + str(job_a_row)].value or "Unknown"
    filename = f"JO&EP - {client}-{length}-{width}-{thickness}-{density}"
    if job_b_row:
        client = source_ws["H" + str(job_b_row)].value or "Unknown"
        length = source_ws["S" + str(job_b_row)].value or "Unknown"
        width = source_ws["T" + str(job_b_row)].value or "Unknown"
        filename += f"-{client}-{length}-{width}"
    filename = sanitize_filename(filename) + ".xlsx"
    filepath = os.path.join(os.path.dirname(__file__), output_dir, filename)
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    counter = 1
    base_filename = filepath
    while os.path.exists(filepath):
        filename = base_filename.replace(".xlsx", f" ({counter}).xlsx")
        filepath = os.path.join(os.path.dirname(__file__), output_dir, filename)
        counter += 1
    template_wb.save(filepath)
    print(f"Saved: {filename}")

# Main Loop Logic
current_a_row = None
current_b_row = None
consecutive_a_rows = []

for row in range(start_row - 1, end_row + 1):
    job_type = source_ws["Q" + str(row)].value
    if job_type == 'A':
        if current_a_row:
            # Calculate the number of pallets and remaining sheets for Job A
            source_value_p = source_ws["P" + str(current_a_row)].value
            source_value_m = source_ws["M" + str(current_a_row)].value
            num_pallets_a = int(source_value_p / source_value_m) if source_value_m else 0
            remaining_sheets_a = source_value_p - num_pallets_a * source_value_m
            
            populate_and_save_template(current_a_row, None, num_pallets_a, remaining_sheets_a)
            consecutive_a_rows.append(current_a_row)
        current_a_row = row

    elif job_type == 'B' and current_a_row:
        # Calculate the number of pallets and remaining sheets for Job B
        source_value_p = source_ws["P" + str(row)].value
        source_value_m = source_ws["M" + str(row)].value
        num_pallets_b = int(source_value_p / source_value_m) if source_value_m else 0
        remaining_sheets_b = source_value_p - num_pallets_b * source_value_m
        
        populate_and_save_template(current_a_row, row, num_pallets_a, remaining_sheets_a, num_pallets_b, remaining_sheets_b)

print("All files created successfully!")
