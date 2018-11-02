Attribute VB_Name = "RunRRForm"

Sub RunRRForm()
    On Error GoTo eh
    
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    RateResetForm.Show

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured in RunRRForm()  " & Err.Description
    GoTo CleanUp
End Sub
