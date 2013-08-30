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
            Range("A1").Select
            Exit For
        End If
    Next
End Sub

Sub FilterOpenOrders()
    Dim PO As String
    Dim ColPO As Integer


    Sheets("Open Orders").Select

    PO = InputBox("Enter the customer PO number", "Enter PO Number")
    ColPO = FindColumn("PO Number")

    If PO = "" Then
        Err.Raise Errors.USER_INTERRUPT, "FilterOpenOrders", "No PO was entered."
    Else
        FilterSheet PO, ColPO, True
    End If

End Sub
