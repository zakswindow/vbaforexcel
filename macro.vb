Sub NextInvoice()
    Range("F8").Value = Range("F8").Value + 1
    Range("A14:F29").ClearContents
    Range("A6:A12").ClearContents
    Range("F33:F34").ClearContents
End Sub
Sub Datacopy()
Worksheets("Invoice").Range("F8").Copy _
        Destination:=Worksheets("Total").Cells(Worksheets("Total").Rows.Count, "C").End(xlUp).Offset(1, 0)
Worksheets("Invoice").Range("F7").Copy _
        Destination:=Worksheets("Total").Cells(Worksheets("Total").Rows.Count, "A").End(xlUp).Offset(1, 0)
Worksheets("Invoice").Range("A6").Copy _
        Destination:=Worksheets("Total").Cells(Worksheets("Total").Rows.Count, "B").End(xlUp).Offset(1, 0)
Worksheets("Invoice").Range("F35").Copy
Sheets("Total").Cells(Rows.Count, "D").End(xlUp).Offset(1, 0).PasteSpecial xlValues
End Sub
Sub SaveBook()
Dim sFile As String
sFile = "MnZ-Invoice" & ".xlsm"
ActiveWorkbook.SaveAs FileName:="/Users/zakhussain/Documents/" & sFile, FileFormat:=52
End Sub
Sub SaveInvWithNewName()
    Dim NewFN As Variant
    ' Copy Invoice to a new workbook
    ActiveSheet.Copy
    NewFN = "/Users/zakhussain/Documents/" & Range("A6").Value & "-" & Range("F8").Value & ".xlsx"
    ActiveWorkbook.SaveAs NewFN, FileFormat:=xlOpenXMLWorkbook
    'Save active workbook as PDF
    'Use this directory format for windows machine
    '"C:\Users\Documents\Invoice\"
    NewFN1 = "/Users/zakhussain/Documents/" & Range("A6").Value & "-" & Range("F8").Value & ".pdf"
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:=NewFN1
    ActiveWorkbook.Close
    Datacopy
    NextInvoice
    ActiveWorkbook.Save
    'ActiveWorkbook.Close
End Sub
