# PythonXmlDocuments
###SOLO APUNTES MANEJO XML PYTHON
##import openpyxl
##
##wb=openpyxl.load_workbook("e.xlsx") #abrimos la hoja de calculos.
##
##print(type(wb))  #comprobamos que esta todo correcto
#canbiar directorios  # os.getcwd()  os.chdir()

#Getting Sheets from the Workbook

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx") #abrimos la hoja de calculos.
##print(wb.sheetnames) # Los nombres de las hojas
##sheet=wb["Hoja3"] # Get a sheet from the workbook
##print(sheet) #<Worksheet "Hoja3">
##print(type(sheet))  #vemos el tipo de objeto
##print(sheet.title)#Get a sheet from the workbook.
##anotherSheet=wb.active  # Get the active sheet
##print(anotherSheet) # imprime la hoja que esta en activo


#Getting Cells from the Sheets

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx") #abrimos como siempre el fichero
##type(wb)
##sheet=wb["Hoja1"] #selecionamos la hoja
##sheet["A1"]# Cogemos la celda de la hoja
##print(sheet["A1"]) # ya tenemos el objeto celda
##print(sheet["A1"].value) # Get the value from the cell
##
##c=sheet["C2"] # Selecionamos otra celda
##print(c.value)  # unprimimos el valor
##
##d=sheet["C1"]  #Selecionamos otra celda
##print(d.value) # imprimimos el valor
##
###Get the row,column ,and value form the cell
##print("Row %s,Column %s is %s" % (c.row, c.column,c.value))#Row 2,Column 3 is 85
##print("Cell %s is %s" % (c.coordinate, c.value)) #Cell C2 is 85
##print(sheet["C1"].value)   #73
##
##
##print(sheet.cell(row=1,column=2))
##print(sheet.cell(row=2,column=1).value)  # da el valor
##
##for i in range(1,8,2): # Go through every other row: # pasar por cada dos filas
##    print(i,sheet.cell(row=i,column=2).value)
##


#####################DETERMINAR EL MAXIMO Y EL MINIMO

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##print(sheet.max_row) #  te indica el maximo largo
##print(sheet.max_column) #te indica el maximo horizontal


############Converting Betwen column Letters and Numbers



##import openpyxl
##from openpyxl.utils import get_column_letter,column_index_from_string
##
##print(get_column_letter(1))   #Convierta la columna 1 en letra
##print(get_column_letter(2))   # Convierte la columna 2 en Letra
##print(get_column_letter(44))  #convierte la columna 44 en letr
##print(get_column_letter(999)) #convierte la columna 999 en letra
##
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##print(get_column_letter(sheet.max_column))  #Te da el maximo fila en este caso C
##column_index_from_string("A")  #Get A`s number.
##column_index_from_string("AA") # Get number AA -> same before but invert




#############Gettin Rows and Columns form the Sheets

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##tuple(sheet["A1":"C3"]) # Coge todas las casillas de la a1 hasta la c3
##print(tuple(sheet["A1":"C3"]))  ##
##for rowOfCellObjects in sheet["A1":"C3"]:
##    for cellObj in rowOfCellObjects:
##        print(cellObj.coordinate,cellObj.value)
##    print("END OF THE ROW")
##
##




