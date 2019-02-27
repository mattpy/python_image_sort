#Please update the variable "address1" and "string_length", as well as update
#the pathname in the final for loop at the end of the script with the correct pathname

import os
import re
import xlsxwriter
from natsort import natsorted, ns

if __name__ == '__main__':
    workbook = xlsxwriter.Workbook('Images.xlsx') # Opens the workbook
    bold = workbook.add_format({'bold':'True'})
    
rows = int(input("Please type in how many columns of images you want (max 6): "))       
bar = float(input("""
Please input the scale for images.
Max value is 0.22
"""))

#update the below variable to include the correct pathname
address1 = os.listdir(r'C:\Users\mccolma5\Desktop\SBImages\\')
#update the above variable to include the correct pathname
address = natsorted(address1)   
print("Here are the addresses", address)

cells = []
#update the below variable to include the correct pathname
string_length = len(os.listdir(r'#your pathname here'))
#update the above variable to include the correct pathname
string_length += 1
print(f"Length of string in directory: {string_length}")

def cell_number_generator(): # Generates cell index values for image placement
    const = 1
    for i in range(1, string_length):
        if rows >= 1:
            var1 = 'A' + str(const)
            cells.append(var1)
        if rows >= 2:
            var1 = 'F' + str(const)
            cells.append(var1)
        if rows >= 3:
            var1 = 'K' + str(const)
            cells.append(var1)
        if rows >= 4:
            var1 = 'P' + str(const)
            cells.append(var1)
        if rows >= 5:
            var1 = 'U' + str(const)
            cells.append(var1)
        if rows >= 6:
            var1 = 'AE' + str(const)
            cells.append(var1)
        const += 13
    return cells

gen1 = cell_number_generator()
print(gen1)

pattern = re.compile(r'\w+')
print('printing gen1', gen1)

#for SB use scale 0.2; for cracking/powdery use .12
worksheet1 = workbook.add_worksheet("Images") # Makes at new worksheet at index[0]

for i in range(1):
    for i, y in zip(gen1, address):
        y1 = y.rsplit('.', 1)[0]
        print("Adding {} to the Excel sheet".format(y1))
        worksheet1.write(i, y1, bold)
        #update the function below with the correct pathname
        worksheet1.insert_image(i, r"your pathname here" + y,
                               {'x_scale': bar, 'y_scale': bar, 'y_offset': 20})
    else:
        print("Creating the workbook, please wait")
    
def close_workbook():
    try:
        workbook.close()
    except PermissionError:
        print(("="*10 + " ")*5,
              "\nData not saved. You must close the workbook before running this script")
    else:
        print("Finished creating workbook, Roll Tide")

close_workbook()
