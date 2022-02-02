# Excel File Merge Script

I made this program quickly with a friend during a data integration project. 
We had to export an excel file with several sheets from Talend. We could have done this directly through Talend, but it seemed easier in python. 
So we developed this script.
Of course we did it very quickly, and it can be greatly improved. It's a V1 in a way.

You will need  Openyxl

#### install openpyxl

````
pip install openpyxl
````

#### More 
Here we have merged only three Excel files into 1. If you want to merge more, add a new loop 

**Exemple :**

```py

#After 
wb3 = xl.load_workbook("/path/file3.xlsx")
ws3 = wb3.worksheets[0]

#Add this
wb4 = xl.load_workbook("/path/file4.xlsx")
ws4 = wb4.worksheets[0]

#After 
wb4.create_sheet('Worksheet3')

# Add this
wb5.create_sheet('Worksheet4')

#After
ws6 = wb5.worksheets[2]
#put another line with an incrementation like
ws7 = wb5.worksheets[3]



# calculate total number of rows and 
# columns in source excel file
mr = ws4.max_row
mc = ws4.max_column
  
# copying the cell values from source 
# excel file to destination excel file
for i in range (1, mr + 1):
    for j in range (1, mc + 1):
        # reading cell value from source excel file
        c = ws4.cell(row = i, column = j)
  
        # writing the read value to destination excel file
        ws7.cell(row = i, column = j).value = c.value
```


Everything is under **MIT licence**, you can re-use this code as much as you like. You can also contribute to its improvement if you wish.