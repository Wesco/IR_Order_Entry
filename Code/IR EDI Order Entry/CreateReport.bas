Attribute VB_Name = "CreateReport"
Option Explicit

Sub GetPO()
    Dim PO As String
    Dim POData As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Integer    ' Column Counter
    Dim j As Long       ' Row Counter
    Dim k As Integer    ' Destination Column

    Sheets("OOR").Select

    PO = InputBox("Enter the PO number", "PO Entry")
    If Trim(PO) <> "" Then
        ActiveSheet.UsedRange.AutoFilter 1, "=" & PO
        ActiveSheet.UsedRange.Copy Destination:=Sheets("PO").Range("A1")

        Sheets("PO").Select
        POData = ActiveSheet.UsedRange
        TotalRows = UBound(POData, 2)
        TotalCols = UBound(POData)
        Cells.Delete

        'Column Order:
        '    1             2                3                   4                     5               6
        'PO Number    Line Number    IR Part Number    IR Part Description    Quantity Ordered    PO Price

        'EDI Column Order:
        '    1          2     3       4       5    6       7        8      9      10      11        12     13     14
        'PO_NUMBER , Branch, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, SHIPTO, NOTE1, NOTE2

        For i = 1 To TotalCols
            For j = 1 To TotalRows
                If i = 1 Then       'PO Number = PO_NUMBER
                    k = 1
                ElseIf i = 2 Then   'Line Number = CUST_LINE
                    k = 4
                ElseIf i = 3 Then   'IR Part Number = PART_NO
                    k = 9
                ElseIf i = 4 Then   'IR Part Description = DESC
                    k = 10
                ElseIf i = 5 Then   'Quantity Ordered = QTY
                    k = 5
                ElseIf i = 6 Then   'PO Price = UNIT_PRICE
                    k = 7
                End If

                Cells(j, k).Value = POData(j, i)
            Next
        Next

    Else
        Err.Raise CustErr.PONOTFOUND, "GetPO", "PO# entry canceled."
    End If
End Sub

Sub CreateOrder()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long

    Sheets("PO").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'EDI Column Order:
    '    A          B     C       D       E    F       G        H      I      J        K         L      M      N
    'PO_NUMBER , Branch, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, SHIPTO, NOTE1, NOTE2

    'Branch
    Range("B1").Value = "Branch"
    Range("B2:B" & TotalRows).Value = "3615"

    'DPC
    Range("C1").Value = "DPC"
    Range("C2:C" & TotalRows).Value = "33454"

    'UOM
    Range("F1").Value = "UOM"
    Range("F2:F" & TotalRows).Formula = "=IFERROR(VLOOKUP(I2,Master!A:D,4,FALSE),"""")"
    Range("F2:F" & TotalRows).Value = Range("F2:F" & TotalRows).Value

    'UNIT_PRICE
    Range("G2:G" & TotalRows).NumberFormat = "$#,##0.00"
    Range("G2:G" & TotalRows).Value = Range("G2:G" & TotalRows).Value

    'SIM
    Range("H1").Value = "SIM"
    Range("H2:H" & TotalRows).Formula = "=IFERROR(VLOOKUP(I2,Master!A:C,3,FALSE),"""")"
    Range("H2:H" & TotalRows).Value = Range("H2:H" & TotalRows).Value

    'DESC
    With Range("J2:J" & TotalRows)
        .Replace ",", ""
        .Replace """", ""
        .Replace ";", ""
    End With
    
    'SHIP_DATE
    Range("K1").Value = "SHIP_DATE"
    
    'SHIPTO
    Range("L1").Value = "SHIPTO"
    Range("L2:L" & TotalRows).Value = "2"

    'NOTE1
    Range("M1").Value = "NOTE1"
    'Range("M2:M" & TotalRows).Formula = "=""LN: "" & D2 & ""   PART: "" & I2"
    'Range("M2:M" & TotalRows).Value = Range("M2:M" & TotalRows).Value
    
    'NOTE2
    Range("N1").Value = "NOTE2"

    'Master Price
    Range("O1").Value = "Master Price"
    Range("O2:O" & TotalRows).Formula = "=IFERROR(VLOOKUP(I2,Master!A:H,8,FALSE),"""")"
    Range("O2:O" & TotalRows).NumberFormat = "$#,##0.00"
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value

    'Make all fonts and borders match
    With ActiveSheet.UsedRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Columns.AutoFit
    End With

    'Highlight pricing discrepancies
    For i = 2 To TotalRows
        If Cells(i, 7).Value <> Cells(i, 15).Value Then
            Range(Cells(i, 1), Cells(i, 15)).Interior.Color = rgbRed
        End If
    Next
End Sub

Sub FormatOOR()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long

    Sheets("OOR").Select

    'Remove report header
    Range("1:3").Delete

    'Remove report footer
    Range(Rows(1).End(xlDown).Row & ":" & ActiveSheet.UsedRange.Rows.Count).Delete

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    For i = TotalCols To 1 Step -1
        If Cells(1, i).Value <> "PO Number" And _
           Cells(1, i).Value <> "Line Number" And _
           Cells(1, i).Value <> "IR Part Number" And _
           Cells(1, i).Value <> "IR Part Description" And _
           Cells(1, i).Value <> "Quantity Ordered" And _
           Cells(1, i).Value <> "PO Price" Then
            Columns(i).Delete
        End If
    Next
End Sub

