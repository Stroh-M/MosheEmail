import csv
# from openpyxl import load_workbook

with open('C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\test.csv', 'r', newline='') as file:
    reader = csv.reader(file)
    rows = list(reader)

    print(rows)

    if "Meir" in rows:
        print("found")

    for x, elem in enumerate(rows):
        if 'Meir' in elem:
            print(x)

            rows[x][1] = 'Rosenberg'

        with open('C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\test.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(rows)


    

        