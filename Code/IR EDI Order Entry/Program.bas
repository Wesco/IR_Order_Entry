Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Main_Error
    
    ImportMaster
    
    MsgBox "Select the 'Supplier Open Order Report'"
    UserImportFile Sheets("OOR").Range("A1")
    FormatOOR
    
    GetPO
    CreateOrder
    
    On Error GoTo 0
    Exit Sub

Main_Error:
    If Err.Number = 18 And Err.Source = "UserImportFile" Then
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
