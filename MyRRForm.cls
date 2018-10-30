VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyRRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Class_Initialize()

End Sub

Private Sub ResetFilters()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub


Sub RunAll(oRR As MyRateReset)

    oRR.RemainingLoanTerm
    oRR.Auth
    oRR.FormatRR
    oRR.JoinData
    oRR.SetMaxExtWithTransScores
    oRR.AdjustMaxExt
    oRR.DeDuplicate
    oRR.RemoveMinMax12
    oRR.FilterRemainingTerm
    oRR.CleanUp

End Sub


Sub ScanCheckboxs(oRR As MyRateReset)

    If RateResetForm.cbCalcRemTerm.Value Then
        oRR.RemainingLoanTerm
    End If
    
    If RateResetForm.cbAuth.Value Then
        oRR.Auth
    End If
    
    If RateResetForm.cbFormatting Then
        oRR.FormatRR
    End If
    
    If RateResetForm.cbJoin Then
        oRR.JoinData
    End If
    
    If RateResetForm.cbMaxExt Then
        oRR.SetMaxExtWithTransScores
    End If
    
    If RateResetForm.cbAdjustMax Then
        oRR.AdjustMaxExt
    End If
    
    If RateResetForm.cbDeduplicate Then
        oRR.DeDuplicate
    End If
    
    If RateResetForm.cbRmMinMax Then
        oRR.RemoveMinMax12
    End If
    
    If RateResetForm.cbFilter Then
        oRR.FilterRemainingTerm
    End If
    
    If RateResetForm.cbClean Then
        oRR.CleanUp
    End If
    
Done:
    Set oRR = Nothing                ' terminate rate reset object
    Exit Sub
eh:
    MsgBox Err.Source & ": bRun_Click() error occured " & Err.Description
    Exit Sub
    
End Sub

Sub CheckAll()
    RateResetForm.cbCalcRemTerm.Value = True
    RateResetForm.cbAdjustMax.Value = True
    RateResetForm.cbAuth.Value = True
    RateResetForm.cbClean.Value = True
    RateResetForm.cbDeduplicate.Value = True
    RateResetForm.cbFilter.Value = True
    RateResetForm.cbJoin.Value = True
    RateResetForm.cbMaxExt.Value = True
    RateResetForm.cbRmMinMax.Value = True
    RateResetForm.cbFormatting.Value = True
End Sub

Sub UncheckAll()
    RateResetForm.cbCalcRemTerm.Value = False
    RateResetForm.cbAdjustMax.Value = False
    RateResetForm.cbAuth.Value = False
    RateResetForm.cbClean.Value = False
    RateResetForm.cbDeduplicate.Value = False
    RateResetForm.cbFilter.Value = False
    RateResetForm.cbJoin.Value = False
    RateResetForm.cbMaxExt.Value = False
    RateResetForm.cbRmMinMax.Value = False
    RateResetForm.cbFormatting.Value = False
End Sub

Sub LogDisable(oRR As MyRateReset)
    RateResetForm.labelBrowse.Visible = False
    RateResetForm.bBrowse.Visible = False
    oRR.LogBool = False
End Sub

Sub LogEnable(oRR As MyRateReset)
    RateResetForm.labelBrowse.Visible = True
    RateResetForm.bBrowse.Visible = True
    oRR.LogBool = True
End Sub

Sub OpenLogFile(sName As String)
    Dim viewLogFile
    viewLogFile = Shell("C:\WINDOWS\notepad.exe """ & sName & """", 1)
End Sub

Sub ChooseLogFile(ByRef sName As String)
    Dim SelectedItem As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select a .txt file to log to"
        .ButtonName = "Select"
        
        If .Show = -1 Then
            'if user clicks OK
            SelectedItem = .SelectedItems(1)
            MsgBox "You have selected: " & SelectedItem
            sName = SelectedItem
            
        Else
        'if user clicks Cancel
        End If
    End With
End Sub


Sub UserFormInit()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    RateResetForm.cbCalcRemTerm.Value = False
    RateResetForm.cbAdjustMax.Value = False
    RateResetForm.cbAuth.Value = False
    RateResetForm.cbClean.Value = False
    RateResetForm.cbDeduplicate.Value = False
    RateResetForm.cbFilter.Value = False
    RateResetForm.cbJoin.Value = False
    RateResetForm.cbMaxExt.Value = False
    RateResetForm.cbRmMinMax.Value = False
    RateResetForm.cbFormatting.Value = False
    
    RateResetForm.obYes.Value = False
    RateResetForm.obNo.Value = True
    RateResetForm.ToggleButton1.Value = False
    
    'RateResetForm.bSelectAll.SetFocus
End Sub

Sub ToggleButton()
    If RateResetForm.ToggleButton1.Value = False Then
    ShowAllOps
    RateResetForm.ToggleButton1.Caption = "Hide"
    Else
    HideAllOps
    RateResetForm.ToggleButton1.Caption = "Show"
    End If
End Sub

Private Sub HideAllOps()
    RateResetForm.cbCalcRemTerm.Visible = False
    RateResetForm.cbAdjustMax.Visible = False
    RateResetForm.cbAuth.Visible = False
    RateResetForm.cbClean.Visible = False
    RateResetForm.cbDeduplicate.Visible = False
    RateResetForm.cbFilter.Visible = False
    RateResetForm.cbJoin.Visible = False
    RateResetForm.cbMaxExt.Visible = False
    RateResetForm.cbRmMinMax.Visible = False
    RateResetForm.cbFormatting.Visible = False
    
    RateResetForm.labelOps.Visible = False
    
    RateResetForm.bSelectAll.Visible = False
    RateResetForm.bDeSelectAll.Visible = False
End Sub

Private Sub ShowAllOps()
    RateResetForm.cbCalcRemTerm.Visible = True
    RateResetForm.cbAdjustMax.Visible = True
    RateResetForm.cbAuth.Visible = True
    RateResetForm.cbClean.Visible = True
    RateResetForm.cbDeduplicate.Visible = True
    RateResetForm.cbFilter.Visible = True
    RateResetForm.cbJoin.Visible = True
    RateResetForm.cbMaxExt.Visible = True
    RateResetForm.cbRmMinMax.Visible = True
    RateResetForm.cbFormatting.Visible = True
    
    RateResetForm.labelOps.Visible = True
    
    RateResetForm.bSelectAll.Visible = True
    RateResetForm.bDeSelectAll.Visible = True
End Sub

'Sub ResetData()
'    Worksheets("DATA_Orig").Range("A1:AH200").Copy
'    Worksheets("Data").Range("A1:AH200").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'    Worksheets("Data").Range("AB1:AB200") = ""
'    Worksheets("Data").Range("AB1").Select
'End Sub

Sub CloseRRForm()
    Unload RateResetForm
    Unload ProgressBar
End Sub


Sub ResetDataSheet()
    Worksheets("Data").Activate
    ResetFilters
    If Sheets("Data").AutoFilterMode Then Cells.AutoFilter
    Worksheets("DATA_Orig").Range("A1:AM200").Copy
    Worksheets("Data").Range("A1:AM200").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Worksheets("Data").Range("AB1:AB200") = ""
    Worksheets("Data").Range("AB1").Select
End Sub




Private Sub Class_Terminate()

End Sub
