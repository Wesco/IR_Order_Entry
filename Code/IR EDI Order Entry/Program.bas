Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.0"
Public Const RepositoryName As String = "IR_Order_Entry"

Sub Main()
    On Error GoTo Main_Error

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ImportGaps
    ImportMaster

    MsgBox "Select the 'Supplier Open Order Report'"
    UserImportFile Sheets("OOR").Range("A1")

    FormatOOR
    GetPO
    CreateOrder
    ExportOrder

    MsgBox "Complete!"
    
    Sheets("Macro").Select
    Range("C7").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    On Error GoTo 0
    Exit Sub

Main_Error:
    If Err.Number = 18 And Err.Source = "UserImportFile" Or Err.Source <> "" Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure " & Err.Source
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module Program"
    End If

End Sub

Sub Clean()
    Dim PrevAlrt As Boolean
    Dim PrevScrn As Boolean
    Dim s As Worksheet

    PrevAlrt = Application.DisplayAlerts
    PrevScrn = Application.ScreenUpdating

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            Cells.Delete
            Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    Application.ScreenUpdating = PrevScrn
    Application.DisplayAlerts = PrevAlrt
End Sub
