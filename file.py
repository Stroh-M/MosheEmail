from openpyxl import load_workbook #type: ignore 

file_path = r'C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\Flat.File.ShippingConfirm (1).xlsx'

# Load workbook
wb = load_workbook(file_path)

# List available sheet names (for debugging)
print("Available sheets:", wb.sheetnames)

# Access specific sheet
sheet = wb['ShippingConfirmation']

# Show existing value before change
print("Before:", sheet.cell(row=2, column=3).value)

# Write new value to C1
sheet.cell(row=2, column=3, value='new value')

# Show value after change
print("After:", sheet.cell(row=1, column=3).value)

# Save the changes
wb.save(file_path)






# import csv
# # from openpyxl import load_workbook

# with open('C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\118545073774020185.txt', 'r') as file:
#     reader = csv.reader(file, delimiter='\t')
#     rows = list(reader)

#     # print(rows)
#     for row in rows:
#         if '114-5883196-2075457' in row:
#             print("found")

#     for x, elem in enumerate(rows):
#         if '114-5883196-2075457' in elem:
#             print(x)

#             rows[x][1] = 'Rosenberg'

#         with open('C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\118545073774020185.txt', 'w', newline='') as file:
#             writer = csv.writer(file, delimiter='\t')
#             writer.writerows(rows)


    

        