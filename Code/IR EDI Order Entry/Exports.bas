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

    'Remove column headers
    Rows(1).Delete

    ActiveSheet.Copy

    ActiveWorkbook.SaveAs FilePath & FileName, xlCSV
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

Sub ExportRemoved()
    Dim PrevDispAlert As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim PONum As String

    Sheets("Removed").Select
    PrevDispAlert = Application.DisplayAlerts
    FilePath = "\\7938-HP02\Shared\IR-Davidson-Mox\Removed Items\"
    FileName = Format(Date, "yyyy-mm-dd") & ".xlsx"
    PONum = Range("A2").Value

    Sheets("Removed").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    Email "nryan@wesco.com", _
          Subject:="IR EDI PO# " & PONum, _
          Body:="The attached list of items needs to be manually entered for IR PO# " & PONum, _
          Attachment:=FilePath & FileName
End Sub
