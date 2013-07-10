Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo ImportFailed
    ImportIR_OOR
    ImportMaster
    On Error GoTo 0
    
    
    CleanOpenOrders

    GoTo ExitSub


ImportFailed:
    Select Case Err.Number
        Case 53:
            MsgBox Err.Description & " could not be found."

        Case Else:
            Err.Raise Err.Number
    End Select
    GoTo ExitSub

ExitSub:
End Sub
