import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

quotation_number = 48
model = 1
path = f"Models/modele{model}.xlsx"
wb = openpyxl.load_workbook(path)

# Params
font = "Roboto"

# Select the active sheet
sheet = wb.active

init_date = "dd/mm/yyyy"

# Add headers for the quotation
headers = ["Item", "Description", "Quantity", "Unit Price", "Total"]
for col_num, header in enumerate(headers, start=1):
    continue

# Add sample data for the quotation
data = [

]

# Populate the data into the sheet
for row_num, row_data in enumerate(data, start=2):
    for col_num, value in enumerate(row_data, start=1):
        continue

B8 = "Caca"

# Get all merged cell ranges
merged_ranges = sheet.merged_cells

# Print merged ranges to verify
for merged_range in merged_ranges:
    print(merged_range)


wb.save(f"Estimates/Devis{quotation_number}.xlsx")

