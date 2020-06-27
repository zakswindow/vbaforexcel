Sub NextInvoice()
    Range("F4").Value = Range("F4").Value + 1
    Range("A14:F29").ClearContents
    Range("A6:A12").ClearContents
End Sub
Sub SaveBook()
Dim sFile As String
sFile = "MnZ-Invoice" & ".xlsm"
ActiveWorkbook.SaveAs Filename:="/Users/zakhussain/Documents/" & sFile, FileFormat:=52
End Sub
Sub SaveInvWithNewName()
    Dim NewFN As Variant
    ' Copy Invoice to a new workbook
    ActiveSheet.Copy
    NewFN = "/Users/zakhussain/Documents/" & Range("A6").Value & "-" & Range("F4").Value & ".xlsx"
    ActiveWorkbook.SaveAs NewFN, FileFormat:=xlOpenXMLWorkbook
    'Save active workbook as PDF
    'Use this directory format for windows machine
    '"C:\Users\Documents\Invoice\"
    NewFN1 = "/Users/zakhussain/Documents/" & Range("A6").Value & "-" & Range("F4").Value & ".pdf"
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=NewFN1
    ActiveWorkbook.Close
    NextInvoice
    ActiveWorkbook.save
    ActiveWorkbook.Close
End Sub
