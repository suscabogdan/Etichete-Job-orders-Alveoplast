import openpyxl
import os
import datetime
import math
from copy import copy
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as openpyxlImage
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.cell import coordinate_from_string
from io import BytesIO
from PIL import Image as PILImage
from PIL import Image, ImageDraw

def get_top_left_cell_of_merged_region(worksheet, cell_address):
    """Identify the top-left cell of a merged region."""
    for merged_range in worksheet.merged_cells.ranges:
        if cell_address in merged_range:
            return worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
    return worksheet[cell_address]

# Ensure the "New Job Orders" directory exists
# output_dir = r"C:\Users\gabri\Desktop\Job Orders Tests\New Job Orders"
output_dir = r"L:\1_EXTRUDARE\01. PLANIFICARE PRODUCTIE\1. CALENDAR COMENZI PRODUCTIE\Script&ETICHETE\Job Orders\New Job Orders"
# output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "New Job Orders")
if not os.path.exists(output_dir):
    os.mkdir(output_dir)

# Load the main data workbook
source_wb = openpyxl.load_workbook("EVIDENTA COMANDA ALVEOPLAST.xlsx")
source_ws = source_wb["COMENZI ALVEOPLAST"]

start_row = int(input("Enter the starting row: "))
end_row = int(input("Enter the ending row: "))

def copy_range(src_ws, dest_ws, src_range, dest_cell):
    copied_images = []  # Create a list to store copied images
    
    rows = src_ws[src_range]
    dest_cell = dest_ws[dest_cell]
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            dest_cell.offset(i, j).value = cell.value
            if cell.has_style:
                dest_cell.offset(i, j)._style = copy(cell._style)

# Create an array of images
images = ["Images/BGMalveoplast Logo.jpeg", "Images/Energie Verde.jpeg", "Images/PP.jpeg", "Images/SARC.jpeg"]
anchor_a = ['AD27', 'U45', 'X23', 'AD46']
anchor_b = ['AP27', 'AG45', 'AJ23', 'AP46']

# Mapping for new Job Oreder Template file
# Mapping for Job B
mapping = {
    "H": "J49", # Denumire client
    "I": "J51", # Cod produs
    "B": "J50", # Numar comanda client
    "P": ["J16", "J53"], # Cantitate de extrudat
    # "X": "J17", # Greutate totala de extrudat
    "T": "J4", # Lungime
    "U": "J5", # Latime
    "W": "J9", # Densitate
    "V": "J7", # Grosime
    "S": "J11", # Culoare
    "M": "J19", # Coli / palet
    # "O": "N24",  # Numar paleti (P/M)
    "R": "L23" # Cod reteta
}

# Mapping for Job A
mapping_job_a = {
    "H": "N49", # Denumire client
    "I": "N51", # Cod produs
    "B": "N50", # Numar comanda client
    "P": ["N16", "N53"], # Cantitate de extrudat
    # "X": "J17", # Greutate totala de extrudat
    "T": "N4", # Lungime
    "U": "N5", # Latime
    "W": "J9", # Densitate
    "V": "J7", # Grosime
    "S": "J11", # Culoare
    "M": "N19", # Coli / palet
    # "O": "N24",  # Numar paleti (P/M)
    "R": "L23" # Cod reteta
}

def load_r_mapping_from_file(filename, color_value):
    r_mapping = {}
    with open(filename, 'r') as file:
        current_key = None
        for line in file:
            line = line.strip()
            if not line:
                continue
            if line.endswith(':'):
                current_key = line[:-1]
                r_mapping[current_key] = {}
            else:
                cell, value = line.split('=', 1)
                # Replace the color placeholder with the actual color value
                value = value.replace('{color}', color_value)
                r_mapping[current_key][cell.strip()] = value.strip()
    return r_mapping

def sanitize_filename(filename):
    """Sanitize filenames."""
    sanitized_name = ''.join(char if char.isalnum() or char in [' ', '_'] else '-' for char in filename)
    return sanitized_name

def populate_and_save_template(job_a_row, job_b_row, num_pallets_a, remaining_sheets_a, num_pallets_b=None, remaining_sheets_b=None):
    if num_pallets_b is None:
        num_pallets_b = 0
    if remaining_sheets_b is None:
        remaining_sheets_b = 0
    
    template_wb = openpyxl.load_workbook('JO&EP Template 2.xlsx')
    template_ws = template_wb["SABLON"]

    for col, dest_cells in mapping_job_a.items():
        if not isinstance(dest_cells, list):
            dest_cells = [dest_cells]
        for dest_cell in dest_cells:
            top_left_cell = get_top_left_cell_of_merged_region(template_ws, dest_cell)
        
            # Default logic
            source_value = source_ws[col + str(job_a_row)].value
            top_left_cell.value = source_value

    if job_b_row:  # Only proceed if there's a B job
        for col, dest_cells in mapping.items():
            if not isinstance(dest_cells, list):
                dest_cells = [dest_cells]
            for dest_cell in dest_cells:
                top_left_cell = get_top_left_cell_of_merged_region(template_ws, dest_cell)
            
                # Default logic
                source_value = source_ws[col + str(job_b_row)].value
                top_left_cell.value = source_value

    # Mapping for values in column R
    color_value = source_ws["S" + str(job_a_row)].value

    r_mapping = load_r_mapping_from_file("Retete.txt", color_value)

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
            src_range = "U1:AE49"
            dest_cell = "U" + str(51 + i * (49 - 16 + 2 + 15))
            copy_range(template_ws, template_ws, src_range, dest_cell)

    # Write the correct information into the Job A labels
    if job_a_row:
        for i in range(num_pallets_a):
            row_v72 = 72 + 15 + i * (49 - 16 + 2 + 15)
            row_v79 = 79 + 15 + i * (49 - 16 + 2 + 15)
            row_AA79 = 79 + 15 + i * (49 - 16 + 2 + 15)

            template_ws.cell(row=row_v72, column=22).value = source_ws.cell(row=current_a_row, column=7).value
            template_ws.cell(row=row_v79, column=22).value = i + 1
            template_ws.cell(row=row_AA79, column=26).value = num_pallets_a

            sheets_per_pallet = int(template_ws['G23'].value)
        
            if i == num_pallets_a - 1:
                remaining_sheets = total_sheets_a - (sheets_per_pallet * (num_pallets_a - 1))
                if remaining_sheets > 0:
                    template_ws.cell(row=row_v72, column=22).value = remaining_sheets
            else:
                template_ws.cell(row=row_v72, column=22).value = sheets_per_pallet


    # Copy and paste the Job B template the correct number of times
    if job_b_row and current_b_row:
        for i in range(num_pallets_b):
            # Copy the Job B template
            src_range = "AG1:AQ49"
            dest_cell = "AG" + str(51 + i * (49 - 16 + 2 + 15))
            copy_range(template_ws, template_ws, src_range, dest_cell)

    # Write the correct information into the Job B labels
    if job_b_row and current_b_row:
        for i in range(num_pallets_b):
            row_ah72 = 72 + 15 + i * (49 - 16 + 2 + 15)
            row_ah79 = 79 + 15 + i * (49 - 16 + 2 + 15)
            row_AL79 = 79 + 15 + i * (49 - 16 + 2 + 15)

            template_ws.cell(row=row_ah72, column=34 - 7).value = source_ws.cell(row=current_b_row, column=14).value
            template_ws.cell(row=row_ah79, column=34).value = i + 1
            template_ws.cell(row=row_AL79, column=38).value = num_pallets_b

            sheets_per_pallet = int(source_ws.cell(row=current_b_row, column=13).value)
        
            if i == num_pallets_b - 1:
                remaining_sheets = total_sheets_b - (sheets_per_pallet * (num_pallets_b - 1))
                if remaining_sheets > 0:
                    template_ws.cell(row=row_ah72, column=34).value = remaining_sheets
            else:
                template_ws.cell(row=row_ah72, column=34).value = sheets_per_pallet
    
    cell_interval = 50

    for i in range(num_pallets_a + 1):
        for j, (img_name, anchor) in enumerate(zip(images, anchor_a)):
            # Compute the anchor for this repetition
            col, row = coordinate_from_string(anchor)
            row += cell_interval * i
            new_anchor = f'{col}{row}'
        
            # Add image
            img = openpyxlImage(img_name)
        
            if j == 0:
                img.width = int(63 * 1.33)
                img.height = int(310 * 1.33)
            elif j == 1:
                img.width = int(138 * 1.33)
                img.height = int(92 * 1.33)
            elif j == 2:
                img.width = int(108 * 1.33)
                img.height = int(109 * 1.33)
            elif j == 3:
                img.width = int(83 * 1.33)
                img.height = int(73 * 1.33)
        
            img.anchor = new_anchor
            template_ws.add_image(img)

        if job_a_row and current_b_row:
            for i in range(num_pallets_b + 1):
                for j, (img_name, anchor) in enumerate(zip(images, anchor_b)):
                    # Compute the anchor for this repetition
                    col, row = coordinate_from_string(anchor)
                    row += cell_interval * i
                    new_anchor = f'{col}{row}'
        
                    # Add image
                    img = openpyxlImage(img_name)
        
                    if j == 0:
                        img.width = int(63 * 1.33)
                        img.height = int(310 * 1.33)
                    elif j == 1:
                        img.width = int(138 * 1.33)
                        img.height = int(92 * 1.33)
                    elif j == 2:
                        img.width = int(108 * 1.33)
                        img.height = int(109 * 1.33)
                    elif j == 3:
                        img.width = int(83 * 1.33)
                        img.height = int(73 * 1.33)
        
                    img.anchor = new_anchor
                    template_ws.add_image(img)



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
    # print(f"Saved: {filename}")
    print(f"Saved {filepath}")

# Main Loop Logic
current_a_row = None
current_b_row = None
consecutive_a_rows = []

for row in range(start_row, end_row + 1):
    job_type = source_ws["Q" + str(row)].value
    current_b_row = row
    if job_type == 'A':
        current_a_row = row
        # Calculate the number of pallets and remaining sheets for Job A
        source_value_p = source_ws["P" + str(current_a_row)].value
        source_value_m = source_ws["M" + str(current_a_row)].value
        num_pallets_a = math.ceil(source_value_p / source_value_m) if source_value_m else 0
        remaining_sheets_a = source_value_p - num_pallets_a * source_value_m
        # Save a separate file for the 'A' job
        populate_and_save_template(current_a_row, None, num_pallets_a, remaining_sheets_a)
    elif job_type == 'B' and current_a_row:
        # Calculate the number of pallets and remaining sheets for Job B
        source_value_p = source_ws["P" + str(row)].value
        source_value_m = source_ws["M" + str(row)].value
        num_pallets_b = math.ceil(source_value_p / source_value_m) if source_value_m else 0
        remaining_sheets_b = source_value_p - num_pallets_b * source_value_m
        # Save a separate file for the 'A' job paired with the 'B' job
        populate_and_save_template(current_a_row, current_b_row, num_pallets_a, remaining_sheets_a, num_pallets_b, remaining_sheets_b)

print("All files created successfully!")
