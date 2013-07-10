Attribute VB_Name = "FormatData"
Option Explicit

Sub CleanOpenOrders()
    Dim TotalRows As Long
    Dim i As Long

    Sheets("Open Orders").Select

    Rows("1:3").Delete
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    For i = TotalRows To 0 Step -1
        If Cells(i, 1).Value = "Grand Total" Then
            Rows(i & ":" & TotalRows).Delete
            Exit For
        End If
    Next


End Sub

