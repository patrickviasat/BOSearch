import pandas as pd
import os.path
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkFont
from datetime import datetime
from openpyxl import Workbook

def searchFile():
    filePath = filedialog.askopenfilename(
        title="Select a file",
        initialdir=defaultDir,
        initialfile=defaultFile,
        filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
    )
    if filePath:
        file_path_label.config(text=f"Selected file: {filePath}")
        
def saveFile(book, jobNum):
    sheet=book.worksheets[0]
    file_path = filedialog.asksaveasfilename(
        initialfile=f"JN{jobNum}_Backorder",
        defaultextension='.xlsx',
        filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
        title="Save the file as..."
    )
    if file_path:
        book.save(file_path)
        print(f"Excel file saved as {file_path}")
    else:
        print("Save operation cancelled")
       
def getComponents():
    if defaultFile == "":
        error1 = tk.Tk()
        error1.title("Error!")
        error1.geometry("300x100")
        Message = tk.Label(error1, text="No file selected!")
        Message.pack(pady=25)
        error1.mainloop()
        return        
        
    jobNum = entry.get()
    if jobNum == "":
        error2 = tk.Tk()
        error2.title("Error!")
        error2.geometry("300x100")
        Message = tk.Label(error2, text="No Job Number input!")
        Message.pack(pady=25)
        error2.mainloop()
        return
    
    file = pd.read_excel(filePath, sheet_name="Original BO report")
    file['Job'] = file['Job'].astype(str)
    results = file[file['Job'] == jobNum]
    
    if results.empty:
        error3 = tk.Tk()
        error3.title("Error!")
        error3.geometry("300x75")
        Message = tk.Label(error3, text="No backordered components with that Job Number!")
        Message.pack(pady=25)
        error3.mainloop()
        return
        
    rowNums = results.index.tolist()
    book = Workbook()
    sheet = book.active
    
    output = tk.Tk()
    output.title("Search Results")
    Title = tk.Label(output, text=f"Job Number: {jobNum}")
    sheet['B1'] = f"Job Number: {jobNum}"
    Title.grid(row=0, column=1)
    Comp = tk.Label(output, text="Component")
    sheet['A2'] = "Component"
    Comp.grid(row = 1, column = 0, padx=10, sticky="w")
    Desc = tk.Label(output, text="Description")
    sheet['B2'] = "Description"
    Desc.grid(row = 1, column = 1, sticky="w")
    Qty = tk.Label(output, text="Quantity Open")
    sheet['C2'] = "Quantity Open"
    Qty.grid(row = 1, column = 2, padx=10, sticky="w")
    i = 2
    seen = set()
    for index in rowNums:
        row = file.loc[index]
        if row['Comp'] not in seen:
            seen.add(row['Comp'])
            CompRes = tk.Label(output, text=row['Comp'])
            sheet[f'A{i + 1}'] = row['Comp']
            CompRes.grid(row = i, column = 0, padx = 10, sticky="w")
            DescRes = tk.Label(output, text=row['Comp Description'])
            sheet[f'B{i + 1}'] = row['Comp Description']
            DescRes.grid(row = i, column = 1, sticky="w")
            QtyRes = tk.Label(output, text=row['Qty Open'])
            sheet[f'C{i + 1}'] = row['Qty Open']
            QtyRes.grid(row = i, column = 2, padx = 10, sticky="w")
            i = i + 1
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 75   
    sheet.column_dimensions['C'].width = 15       
    saveButton = tk.Button(output, text="Save as Excel File", command = lambda: saveFile(book, jobNum))
    saveButton.grid(row = i, column = 1, pady = 10)
    output.mainloop()
    
now = datetime.now()
year = now.year
yearEnd = "{:02d}".format(year % 100)
fileName = now.strftime("%m%d" + yearEnd)

if now.month >= 4:
    fiscalYear = int(yearEnd) + 1
    fiscalYear = str(fiscalYear)
else:
    fiscalYear = yearEnd

defaultDir = "None Selected "
defaultFile = ""
if os.path.exists(f"//dc1nas/Projects-AB/AS/Data-OPS/80_Operations_Data_(Leka)/30_Inventory/30_Inventory/Backorder Metrics/FY{str(fiscalYear)} Updated Daily Reports/{fileName}.xlsm"):
    defaultDir = f"//dc1nas/Projects-AB/AS/Data-OPS/80_Operations_Data_(Leka)/30_Inventory/30_Inventory/Backorder Metrics/FY{str(fiscalYear)} Updated Daily Reports"
    defaultFile = f"{fileName}.xlsm"
    
filePath = defaultDir + "/" + defaultFile    
    
root = tk.Tk()
root.geometry('1100x150')
root.title("Backorder Component Search")

frame1 = tk.Frame(root)
frame2 = tk.Frame(root)

default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(weight="bold")

fileLabel = tk.Label(root, text="Update to today's backorder report (if necessary):")
fileLabel.pack()

#Search File GUI
file_path_label = tk.Label(frame2, text=f"Selected file: {filePath}")
file_path_label.pack(side=tk.LEFT)
searchFileButton = tk.Button(frame2, text="Browse...", command=searchFile)
searchFileButton.pack(side=tk.LEFT)
frame2.pack()

#JN Entry GUI
entryLabel = tk.Label(frame1, text="Enter Job Number:")
entryLabel.pack(side=tk.LEFT)
entry = tk.Entry(frame1)
entry.pack(pady=10, side=tk.LEFT)
frame1.pack()

#Components Button
getComponentsButton = tk.Button(root, text="Get Components", command=getComponents)
getComponentsButton.pack(pady=10)
root.bind('<Return>', lambda event: getComponents())
root.mainloop()