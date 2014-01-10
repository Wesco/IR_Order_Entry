Attribute VB_Name = "Imports"
Option Explicit

Sub ImportMaster()
    Dim FileName As String
    Dim FilePath As String

    FileName = "IR Master " & Format(Date, "yyyy") & ".xlsx"
    FilePath = "\\br3615gaps\gaps\IR\Master\"

    Workbooks.Open FilePath & FileName
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
    ActiveWorkbook.Close
End Sub
