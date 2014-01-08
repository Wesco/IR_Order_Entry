Attribute VB_Name = "Exports"
Option Explicit

Sub ExportOrder()
    Dim PrevDispAlert As Boolean
    Dim FilePath As String
    Dim FileName As String
    
    Sheets("PO").Select
    PrevDispAlert = Application.DisplayAlerts
    FilePath = "\\idxexchange-new\EDI\Spreadsheet_PO\"
    FileName = Range("A2").Value
    
    
    'Remove "Master Price"
    Columns("O:O").Delete
    
    'Remove column headers
    Rows(1).Delete
    
    ActiveSheet.Copy
    
    ActiveWorkbook.SaveAs FilePath & FileName, xlCSV
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub
