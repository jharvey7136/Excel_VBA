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
    CheckListBoxes
    Dim oProgBar As New MyProgressBar   'init new progress bar
    
    oProgBar.UpdateProgressBar 1, 100, , "Calculating Remaining Term.."
    oRR.RemainingLoanTerm
    oProgBar.UpdateProgressBar 10, 100, , "Generating authorization code.."
    oRR.Auth
    oProgBar.UpdateProgressBar 20, 100, , "Applying formatting.."
    oRR.FormatRR
    oProgBar.UpdateProgressBar 30, 100, , "Joining data with trans union.."
    
    oRR.JoinData RateResetForm.sTransSheet, RateResetForm.sDataSheetName
    
    oProgBar.UpdateProgressBar 40, 100, , "Calculating maximum extension.."
    oRR.SetMaxExtWithTransScores
    oProgBar.UpdateProgressBar 50, 100, , "Adjusting Max Extension.."
    oRR.AdjustMaxExt
    oProgBar.UpdateProgressBar 60, 100, , "Removing duplicates.."
    oRR.DeDuplicate
    oProgBar.UpdateProgressBar 70, 100, , "Removing Min/Max = 12.."
    oRR.RemoveMinMax12
    oProgBar.UpdateProgressBar 80, 100, , "Filtering remaining term.."
    oRR.FilterRemainingTerm
    oProgBar.UpdateProgressBar 90, 100, , "Cleaning sheet.."
    oRR.CleanUp
    oProgBar.UpdateProgressBar 100, 100, , "Operations Complete!"
    
    oRR.EndingAnalysis

End Sub


Sub ScanCheckboxs(oRR As MyRateReset)
    CheckListBoxes
    Dim oProgBar As New MyProgressBar

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
        oRR.JoinData RateResetForm.sTransSheet, RateResetForm.sDataSheetName
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
    On Error GoTo eh
    Dim viewLogFile
    viewLogFile = Shell("C:\WINDOWS\notepad.exe """ & sName & """", 1)
Done:
    Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured in MyRRForm.OpenLogFile(): " & Err.Description
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
    
    RateResetForm.bRun.Enabled = False
    RateResetForm.bRunAll.Enabled = False
    RateResetForm.obFirstHalf.Value = True
    
    RateResetForm.lFilePath.Caption = "FilePath: " & Application.ActiveWorkbook.FullName
    
    For n = 1 To ActiveWorkbook.Sheets.Count        'fill list boxes with sheet names
        With RateResetForm.ListBox1
            .AddItem ActiveWorkbook.Sheets(n).Name
        End With
        With RateResetForm.ListBox2
            .AddItem ActiveWorkbook.Sheets(n).Name
        End With
        With RateResetForm.ListBox3
            .AddItem ActiveWorkbook.Sheets(n).Name
        End With
    Next n
    
    For n = 1 To 12         'fill list box with month names
        With RateResetForm.ListBox4
            .AddItem MonthName(n)
        End With
    Next n
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

Sub CloseRRForm()
    Unload RateResetForm
    Unload ProgressBar
End Sub


Sub ResetDataSheet()
    On Error GoTo eh
    
    If RateResetForm.sDataSheetName = "" Then
        MsgBox "Select the working sheet you wish to reset back to the state of the original data sheet"
        RateResetForm.ListBox1.SetFocus
        Exit Sub
    End If
    If RateResetForm.sDataOrigName = "" Then
        MsgBox "Select the original data sheet from the list box to reset to"
        RateResetForm.ListBox2.SetFocus
        Exit Sub
    End If

    Worksheets(RateResetForm.sDataSheetName).Activate
    ResetFilters
    
    If Sheets(RateResetForm.sDataSheetName).AutoFilterMode Then Cells.AutoFilter
    Worksheets(RateResetForm.sDataOrigName).Range("A1:AM200").Copy
    Worksheets(RateResetForm.sDataSheetName).Range("A1:AM200").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Worksheets(RateResetForm.sDataSheetName).Range("AB1:AB200") = ""
    Worksheets(RateResetForm.sDataSheetName).Range("AB1").Select
    
Done:
    Exit Sub
eh:
    MsgBox Err.Source & ": Error occured " & Err.Description
    Exit Sub
End Sub


Sub CheckListBoxes()
    If RateResetForm.sTransSheet = "" Then
        MsgBox "Select the trans union sheet"
        RateResetForm.ListBox3.SetFocus
        Exit Sub
    End If
    If RateResetForm.sDataSheetName = "" Then
        MsgBox "Select the working data sheet"
        RateResetForm.ListBox1.SetFocus
        Exit Sub
    End If
End Sub


Sub FirstHalf(oRR As MyRateReset)
    oRR.bFirstHalf = True
End Sub

Sub SecondHalf(oRR As MyRateReset)
    oRR.bFirstHalf = False
End Sub


Sub CheckRunReqs()
    With RateResetForm
        If .sDataSheetName <> "" And _
           .sTransSheet <> "" And _
           .sMonthName <> "" Then
           
           .bRunAll.Enabled = True
        End If
    End With
End Sub


Sub CheckRunSelected()  'check if any of the checkboxes are selected

    With RateResetForm
        If .cbAdjustMax Or _
            .cbAuth Or _
            .cbCalcRemTerm Or _
            .cbClean Or _
            .cbDeduplicate Or _
            .cbFilter Or _
            .cbFormatting Or _
            .cbJoin Or _
            .cbMaxExt Or _
            .cbRmMinMax Then
            
            .bRun.Enabled = True
        Else
            .bRun.Enabled = False
        End If
    End With
End Sub



Private Sub Class_Terminate()

End Sub
