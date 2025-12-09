- note -

i have created a folder named "INVOICE" in local disk E. so for saving files we should have a folder in local disk E or we can change the file saving path by replacing 
"E:\INVOICE\" to any location.


functions and features used:

vlookup - for getting values of hsn code from stock page.
data validation - to insert data from available stock.
max - used to increment invoice number.
macros and VBA - to print the created invoice, to save sales entries by assigning macros to shapes.
sum - grand total
absolute and relative references - for auto increment of serial numbers 

---
the vba code for printing and saving files
-

Sub saveinvoice()
Dim saverng As Range
Dim pdfname As String
Dim path As String

Set saverng = Range("A1:F35")
pdfname = Range("B8") & "" & Range("F3")
path = "E:\INVOICE\"
saverng.ExportAsFixedFormat xITypePDF, Filename:=path & pdfname
End Sub

--
