Sub NextInvoice()
    Range("F4").Value = Range("F4").Value + 1
    Range("A14:F29").ClearContents
    Range("A6:A12").ClearContents
End Sub
Sub SaveInvWithNewName()
    Dim NewFN As Variant
    ' Copy Invoice to a new workbook
    ActiveSheet.Copy
    NewFN = "/Users/<username>/Documents/" & Range("A6").Value & "-" & Range("F4").Value & ".xlsx"
    ActiveWorkbook.SaveAs NewFN, FileFormat:=xlOpenXMLWorkbook
    'Save active workbook as PDF
    'Use this directory format for windows machine
    '"C:\Users\Documents\Invoice\"
    NewFN1 = "/Users/<username>/Documents/" & Range("A6").Value & "-" & Range("F4").Value & ".pdf"
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=NewFN1
    ActiveWorkbook.Close
    NextInvoice
    ActiveWorkbook.save
    ActiveWorkbook.Close
End Sub
