Attribute VB_Name = "Imports"
Option Explicit

Sub ImportMaster()
    Const Path As String = "\\br3615gaps\gaps\IR\Master\"
    Dim FileName As String
    Dim PrevDispAlert As Boolean

    FileName = "IR Master " & Format(Date, "yyyy") & ".xlsx"
    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    If FileExists(Path & FileName) Then
        Workbooks.Open Path & FileName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
        ActiveWorkbook.Close
    Else
        Application.DisplayAlerts = PrevDispAlert
        Err.Raise 53, Description:="IR Master"
    End If

    Application.DisplayAlerts = PrevDispAlert
End Sub

Sub ImportIR_OOR()
    Const Path As String = "\\7938-HP02\Shared\IR order entry\IR macro for all plant order entry\IR Open Purchase Orders\"
    Dim FileName As String
    Dim PrevDispAlert As Boolean
    
    PrevDispAlert = Application.DisplayAlerts
    FileName = "Open POs" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    Application.DisplayAlerts = False

    If FileExists(Path & FileName) Then
        Workbooks.Open Path & FileName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Open Orders").Range("A1")
        ActiveWorkbook.Close
    Else
        Application.DisplayAlerts = PrevDispAlert
        Err.Raise 53, Description:="Open POs " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    End If
    
    Application.DisplayAlerts = PrevDispAlert
End Sub
