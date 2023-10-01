import re
import csv
import pandas as pd
import xlwings as xw
import time
import os

# User to change this
prt_file = r'./BP24_PP_HM23V3_6IC_3IW_5TL_137_RMS1HU_R3_unconstrainedwater_IPR.PRT'
xlsx_file = f'PI_calculated.xlsx'

# User not to change this
start_time = time.time()
csv_file = f'output_IPR.csv'
IPR_row_count = 0

#########################################################
# Step 1: Extract IPR tables and put them in a CSV file
#########################################################

def extract_table_data(text):
    lines = text.split('\n')
    data = []

    for line in lines:
        if '|' in line:
            row = [cell.strip() for cell in line.split('|') if cell.strip()]
            data.append(row)

    return data

def write_header_to_csv(csv_file):
    with open(csv_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Date', 'Well Number', 'BOTTOM_HOLE_PRESSURE (bar)', 'OIL_PRODUCTION_RATE (sm3/d)', 'GAS_PRODUCTION_RATE (sm3/d)', 'WATER_PRODUCTION_RATE (sm3/d)'])

def write_data_to_csv(well_number, table_data, csv_file):
    with open(csv_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        for row in table_data:
            writer.writerow([well_number] + row)

def backfill_date(csv_file, date):
    with open(csv_file, mode='r', newline='') as file:
        reader = csv.reader(file)
        rows = list(reader)

        updated_rows = []
        for row in rows:
            if len(row) == 5: # well, BHP, oil, gas, water
                updated_row = [date] + row
                updated_rows.append(updated_row)
            else:
                updated_rows.append(row)

    with open(csv_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(updated_rows)

with open(prt_file, 'r') as file:
    content = file.read()

    # Split the input text into lines
    lines = content.split('\n')

    # Initialize a flag to indicate if we are inside an IPR table
    inside_table = False

    # Initialize a variable to store the current well number
    current_well_number = None

    # Initialize a list to store the lines of the current table
    current_table = []

    write_header_to_csv(csv_file)

    for line in lines:
        if "REPORT   IPR table for well Well:" in line:
            # If we encounter the start of an IPR table, extract the well number
            current_well_number = re.search(r'Well:(\S+)', line).group(1)
            inside_table = True
            IPR_row_count = 0
            continue
        if inside_table and 'GAS_INJECTION_RATE' in line:
            inside_table = False
        if inside_table and '|' in line and not '|-' in line and not 'BOTTOM_HOLE_PRESSURE' in line:
            # If we are inside an IPR table and the line contains '|', it's part of the table
            current_table.append(line)
            # Count the number of rows in the IPR table
            IPR_row_count += 1
        if inside_table and not line.strip():
            # If we encounter an empty line, it indicates the end of the table
            table_text = '\n'.join(current_table)
            table_data = extract_table_data(table_text)
            write_data_to_csv(current_well_number, table_data, csv_file)
            print(f'Successfully extracted and saved data for well {current_well_number} to {csv_file}.')
            inside_table = False
            current_table = []
        elif "SECTION  The simulation has reached" in line:
            date = re.search(r'reached (\S+)', line).group(1)
            print(f'Date is: ', date)
            backfill_date(csv_file, date)

end_time = time.time()
elapsed_time = int(end_time - start_time)
print(f'IPR extraction completed in {elapsed_time} seconds')

#########################################################
# Step 2: Convert CSV file to Excel and calculate PI
#########################################################

# Load the CSV file into a pandas DataFrame
df = pd.read_csv(csv_file)

# Save the DataFrame as an Excel file
df.to_excel(xlsx_file, index=False)

# Load the Excel file
wb = xw.Book(xlsx_file)
sheet = wb.sheets[0]
wb.sheets[0].name = 'Calculate PI'

sheet.range(1,7).value = 'OIL_PI (sm3/d.bar)'
sheet.range(1,8).value = 'GAS_PI (sm3/d.bar)'
sheet.range(1,9).value = 'WATER_PI (sm3/d.bar)'

# Create new sheet
wb.sheets.add('PI summary')
paste_sheet = wb.sheets['PI summary']

paste_sheet.range(1,1).value = 'Date'
paste_sheet.range(1,2).value = 'Well'
paste_sheet.range(1,3).value = 'SBHP (bara)'
paste_sheet.range(1,4).value = 'Oil_PI (sm3/d.bar)'
paste_sheet.range(1,5).value = 'Gas_PI (sm3/d.bar)'
paste_sheet.range(1,6).value = 'Water_PI (sm3/d.bar)'

IPR_rows_taken = str(IPR_row_count - 1)
paste_row = 1

# Iterate through the rows
for row in sheet.range('A1').expand('down'):
    # Do the following in increments of the size of the IPR table, at the last row of each table
    if row.row > 1 and (row.row - 1) % 5 == 0:
        paste_row += 1
        # Calculate the slope using the formula
        formula = '=-SLOPE(INDEX($C:$F,ROW()-' + IPR_rows_taken + ',COLUMN()-5):INDEX($C:$F,ROW()-1,COLUMN()-5),INDEX($C:$F,ROW()-' + IPR_rows_taken + ',1):INDEX($C:$F,ROW()-1,1))'
        row.offset(0,6).formula = formula
        row.offset(0,7).formula = formula
        row.offset(0,8).formula = formula
        date = row.offset(0,0).value
        well = row.offset(0,1).value
        sbhp = row.offset(0,2).value
        oil_pi = row.offset(0,6).value
        gas_pi = row.offset(0,7).value
        water_pi = row.offset(0,8).value
        paste_sheet.range(paste_row,1).value = date
        paste_sheet.range(paste_row,2).value = well
        paste_sheet.range(paste_row,3).value = sbhp
        paste_sheet.range(paste_row,4).value = oil_pi
        paste_sheet.range(paste_row,5).value = gas_pi
        paste_sheet.range(paste_row,6).value = water_pi
        try:
            well = int(well)
        except:
            pass
        print(f'Calculated PI for well {well} at {date}')

# Save the modified Excel file
wb.save(xlsx_file)
wb.close()
os.remove(csv_file)

print(f'{xlsx_file} saved and closed.')

end_time = time.time()
elapsed_time = end_time - start_time - elapsed_time
minutes = int(elapsed_time // 60)
seconds = int(elapsed_time % 60)
print(f'PI calculation and compilation completed in {minutes} minutes and {seconds} seconds')