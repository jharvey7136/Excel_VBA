Attribute VB_Name = "ReturnedChecks"
'
'
'
'
'

Sub RunReturnedChecks()
    On Error GoTo eh                            ' error handling
    Application.Calculation = xlAutomatic       ' turn on auto calculations if previously turned off
    
    Dim oReturnedChecks As New MyReturnedChecks ' initialize new returned checks object
    
    oReturnedChecks.Formatting                  ' apply formatting
    oReturnedChecks.RetReasonReplace            ' replace ret cd with description
    oReturnedChecks.TurnFeePositive             ' turn negatives to positives
    oReturnedChecks.MaskAccountNumbers          ' mask off all but last four digits of acct number
    oReturnedChecks.FilterAndDelete             ' remove extra acct numbers
    oReturnedChecks.Resize                      ' autofit all columns
    oReturnedChecks.SaveFile                    ' save file

Done:
    Exit Sub
eh:
    MsgBox Err.Source & ": RunRateReset() error occured " & Err.Description
    Exit Sub
End Sub
