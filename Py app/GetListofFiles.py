import glob
import os
from openpyxl import load_workbook

mark_dir = "C:\\Users\\DmitryDmytrenko\\Documents\\Technology\\DD\\Zendesk\\Py app\\Files From Brad Mark\\*.xlsx"
wb = load_workbook(filename=max(glob.glob(mark_dir), key=os.path.getctime))
sheet = wb.worksheets[0]
print(max(glob.glob(mark_dir), key=os.path.getctime))
print(sheet)

reply = int(input("Please enter 1 to continue with the file and tab above "))

if reply == 1:
    print("Will process the file above to create ZD tickets!")

else:
    print("Tickets won't be created")
    exit()

print(max(glob.glob(mark_dir), key=os.path.getctime))
print(wb)