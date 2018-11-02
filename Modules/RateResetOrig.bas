Attribute VB_Name = "RateResetOrig"
'                                                                                                                  '
' -------------------------------------------------- RATE RESET -------------------------------------------------- '
'   This module does the heavy lifting in the rate reset process of preparing the                                  '
' upload file. With the press of a button, all of the formatting and calculating                                   '
' is taken care of without having to do anything besides press a button.                                           '
'                                                                                                                  '
' What the module does:                                                                                            '
'   Remaining Loan Term: calculates the amount of months left on the loan                                          '
'   Generate Auth code: building authoriziation code using SSN and birthday                                        '
'   Formatting: sets all the correct formatting to meet the upload specifications                                  '
'   Set Max Extension: uses the credit score to calculate the maximum load extension in months                     '
'   Adjust Max Extension: adjusts max extension if exceeds 3 years out from current term with a max of 84 months   '
'   Deduplicate: Removes duplicates                                                                                '
'   Removes Min/Max if both equal 12: Cannot have min and max extension both equal 12 months                       '
'   Filter remaining term: only want loans with >13 months remaining                                               '
'   Clean up: pastes final version onto a new sheet                                                                '
'   Run all commands: one subroutine to run all everything in the correct order                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' initialize new Logger
Dim oLogger As New MyLogger


Sub ResetFilters()
'
' Function to clear all filter. Does not remove filters, just shows all data
'
    On Error Resume Next
    ActiveSheet.ShowAllData
    
End Sub

Sub RemainingLoanTerm()
'
' RemainingLoanTerm Macro
' Function to calculate the Remaining Loan Term in months
'
    oLogger.Logger "Subroutine", "RemainingLoanTerm()", "Calculating the Remaining Loan Term in months.."
    
    On Error GoTo eh

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row

' Compute (prod date - dateOfNote) = number of days loan has been active
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "days"
    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-12]"
    Range("AC2").Select
    Range("AC2").AutoFill Destination:=Range("AC2:AC" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "RemainingLoanTerm()", "Number of days loan has been active... Done"
    
' Convert days to months = (days active * 0.032855) = number of months loan has been active
    Range("AD1").Select
    ActiveCell.FormulaR1C1 = "months"
    Range("AD2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*0.032855"
    Range("AD2").Select
    Range("AD2").AutoFill Destination:=Range("AD2:AD" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "RemainingLoanTerm()", "Days active converted to months... Done"
    
' Compute remaining term = (original term in months - months loan has been active) = remaining loan term in months
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "remaining"
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]-RC[-1]"
    Range("AE2").Select
    Range("AE2").AutoFill Destination:=Range("AE2:AE" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "RemainingLoanTerm()", "Remaining loan term in months... Done"
    
' Paste into column K: RemainingLoanTerm
    Range("AE2:AE" & LastRow).Select
    Selection.Copy
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K1").Select
    
Done:
    oLogger.Logger "Done", "RemainingLoanTerm()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "RemainingLoanTerm()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in RemainingLoanTerm(): " & Err.Description
    
End Sub

Sub Auth()
'
' Auth Macro
' Function to create the Auth1 code using the year of birth and SSN
'
    oLogger.Logger "Subroutine", "Auth()", "Creating the Auth1 code using year of birth and last four SSN..."

    On Error GoTo eh

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row
    
' Fetch last 4 digits of SSN
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "last 4"
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-8],4)+0"
    Range("AG2").Select
    Range("AG2").AutoFill Destination:=Range("AG2:AG" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "Auth()", "Fetching last four digits of SSN... Done"
    
' Concatenate year of birth and last four of SSN: 1985 + 3746 = Auth1 = 19853746
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "auth"
    Range("AH2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-8],TEXT(RC[-1],""0000""))"
    Range("AH2").Select
    Range("AH2").AutoFill Destination:=Range("AH2:AH" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "Auth()", "Concatinating year of birth and last four SSN... Done"
    
' Paste into column B: Auth1
    Range("AH2:AH" & LastRow).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    
Done:
    oLogger.Logger "Done", "Auth()", "Subrouting finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "Auth()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub FormatRR()
'
' FormatRR Macro
' Fix certain formatting
'
    oLogger.Logger "Subroutine", "FormatRR()", "Applying proper formatting..."

' LoanNumber -> Number with no decimals
    Columns("I:I").Select
    Selection.NumberFormat = "0"
    
' RemainingLoanTerm -> Number with no decimals
    Columns("K:K").Select
    Selection.NumberFormat = "0"
    
' CurrentPrincipalBalance -> Number with two decimal places
    Columns("R:R").Select
    Selection.NumberFormat = "0.00"
    
' CurrentIntRate -> Number with two decimal places
    Columns("S:S").Select
    Selection.NumberFormat = "0.00"
    
' CurrentPIPayment -> Number with two decimal places
    Columns("T:T").Select
    Selection.NumberFormat = "0.00"
    
' OriginalLoanAmount -> Number with two decimal places
    Columns("W:W").Select
    Selection.NumberFormat = "0.00"
    
    oLogger.Logger "Done", "FormatRR()", "Subroutine finished"
    
End Sub

Sub FilterRemainingTerm()
'
' FilterRemainingTerm Macro
' Only want loans with >13 months remaining
'
    oLogger.Logger "Subroutine", "FilterRemainingTerm()", "Fetching loans with >13 months remaining..."
    On Error GoTo eh

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row

' Activate filter on column K: RemainingLoanTerm
    Range("K1").Select
    Selection.AutoFilter
    
' Filter out only loans with greater than 13 months remaining
    ActiveSheet.Range("$A$1:$AE$" & LastRow).AutoFilter Field:=11, Criteria1:=">13", _
        Operator:=xlAnd
        
Done:
    oLogger.Logger "Done", "FilterRemainingTerm()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "FilterRemainingTerm()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    
End Sub
Sub CleanUp()
'
' Cleanup Macro
' remove un-needed columns, move filtered data to new sheet
'
    oLogger.Logger "Subroutine", "Cleanup()", "Removing un-needed data, moving filtered data to new sheet..."
    On Error GoTo eh

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row

' Copy columns A through W (only columns needed)
    Range("A1:W" & LastRow).Select
    Selection.Copy
    
' Paste into a new sheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    
' Resize all columns to AutoFit
    Columns("A:A").EntireColumn.AutoFit
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
Done:
    oLogger.Logger "Done", "Cleanup()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "Cleanup()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description

End Sub
Sub AdjustMaxExt()
'
' AdjustMaxExt Macro
' Loan term not to exceed more than 3 years out from current term with a max of 84 months
' Test if RemainingLoanTerm + LoanExtensionMax > 84
'   if true, then LoanExtensionMax = 84 - RemainingLoanTerm
    
    oLogger.Logger "Subroutine", "AdjustMaxExt()", "Adjusting max extensions if necessary..."
    On Error GoTo eh

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row
    
' Supply =IF formula to test
    Range("AJ2").Select
    ActiveCell.FormulaR1C1 = "=IF((RC[-25]+RC[-23])>84,(84-RC[-25]),RC[-23])"
    Range("AJ2").Select
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & LastRow), Type:=xlFillDefault
    oLogger.Logger "Subroutine", "AdjustMaxExt()", "Testing: If RemainingLoanTerm + LoanExtensionMax > 84... Done"
    
' Copy and Paste new MaxExt values into column M
    Range("AJ2:AJ" & LastRow).Select
    Selection.Copy
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    oLogger.Logger "Subroutine", "AdjustMaxExt()", "Applying new MaxExtension results... Done"
    
' Format values -> Number -> no decimal places
    Selection.NumberFormat = "0"
    Range("M1").Select
    
Done:
    oLogger.Logger "Done", "AdjustMaxExt()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "AdjustMaxExt()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub DeDuplicate()
'
' DeDuplicate Macro
' Scan sheet for duplicates, exclude LoanExtMax because some people may have multiple credit scores
' resulting in mulitple lines for the same loan. Excluding this column will remove any dupes.
'
    On Error GoTo eh
    oLogger.Logger "Subroutine", "DeDuplicate()", "Scanning sheet for duplicates..."

    LastRow = Cells(Rows.Count, "AA").End(xlUp).Row

' Execute dupe removal
    Range("I1").Select
    ActiveSheet.Range("$A$1:$AJ$" & LastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, _
        7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27), Header:=xlYes
        
Done:
    oLogger.Logger "Done", "DeDuplicate()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    oLogger.Logger "Error", "DeDuplicate()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description

End Sub
Sub RemoveMinMax12()
'
' Subroutine to remove rows where the min and max extension both equal 12
'
    On Error GoTo eh
    oLogger.Logger "Subroutine", "RemoveMinMax12()", "Scanning sheet for min ext and max ext both equal to 12..."

    ' Determine number of rows (filtered)
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Filter LoanExtMin by "12"
    ActiveSheet.Range("$A$1:$AA$" & LastRow).AutoFilter Field:=12, Criteria1:="=12"

    ' Filter LoanExtMax by "12"
    ActiveSheet.Range("$A$1:$AA$" & LastRow).AutoFilter Field:=13, Criteria1:="=12"

    ' Delete only visible (filtered) rows
    ActiveSheet.Range("$A$1:$AA$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    
    ResetFilters
Done:
    oLogger.Logger "Done", "RemoveMinMax12()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger "Error", "RemoveMinMax12()", "The following error occured: " & Err.Description
End Sub

Sub SetMaxExtWithTransScores()
'
' Use the new credit score added with the trans union file to determine the max
' amount of months the loan can be extened
'
    oLogger.Logger "Subroutine", "SetMaxExtWithTransScores()", "Using new credit scores to re-calculate max extension..."
    On Error GoTo eh
    
    ' variables
    Dim myDataRng As Range
    Dim cell As Range
    Dim src As Integer

    ' set the range
    Set myDataRng = Range("AB2:AB" & Cells(Rows.Count, "AA").End(xlUp).Row)
    
    ' loop through range of cells, perform case switch to get max extension in months
    For Each cell In myDataRng
        If Not IsEmpty(cell) Then
            src = cell.Value
            cell.Value = CaseSwitch(src)
        End If
    Next cell
    oLogger.Logger "Subroutine", "SetMaxExtWithTransScores()", "Case switch complete..."
    
    ' Paste into column M: LoanExtensionMax, skip blanks
    Range("AB2:AB" & Cells(Rows.Count, "AA").End(xlUp).Row).Select
    Selection.Copy
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
        
Done:
    oLogger.Logger "Done", "SetMaxExtWithTransScores()", "Subroutine finished"
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger "Error", "SetMaxExtWithTransScore()", "The following error occured: " & Err.Description

End Sub

Function CaseSwitch(ByRef score As Integer)
'
' function that takes the credit score from the trans union file column and returns
' 36, 24, or 12 depending on the credit score
'
    oLogger.Logger "Function", "caseSwitch(ByRef score As Integer)", "Performing case switch on trans scores..."
    On Error GoTo eh

    Dim result As Integer
    
    Select Case score
    
        Case 650 To 1000
            result = 36
        
        Case 600 To 649
            result = 24
            
        Case 0 To 599
            result = 12
            
    End Select
    
    CaseSwitch = result
    
Done:
    oLogger.Logger "Done", "caseSwitch(ByRef score As Integer)", "Function finished"
    Exit Function
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger "Error", "caseSwitch(ByRef score As Integer)", "The following error occured: " & Err.Description

End Function

Sub RunAllCommands()
'
' RunAllCommands Macro
' Will execute all the previous functions in sequential order
'
    oLogger.Logger "Subroutine", "RunAllCommands()", "RateReset beginning to run all commands..."

    On Error GoTo eh

    Application.Calculation = xlAutomatic

    RemainingLoanTerm               ' (prod date - dateOfNote)
    Auth                            ' year of birth + last four of SSN: 1985 + 3746 = Auth1 = 19853746
    FormatRR                        ' number formatting
    SetMaxExtWithTransScores        ' use credit scores from trans union file to update max extension
    AdjustMaxExt                    ' if RemainingLoanTerm + LoanExtensionMax > 84, then LoanExtensionMax = 84 - RemainingLoanTerm
    DeDuplicate                     ' remove duplicate members due to possible multiple credit scores
    RemoveMinMax12                  ' remove rows where min and max extension both equal 12
    FilterRemainingTerm             ' only want loans with >13 months remaining
    CleanUp                         ' move final data to a new sheet

Done:
    ' if process is successful, show msgbox with number of rows
    Dim k As Long
    k = ActiveSheet.Range("A1048576").End(xlUp).Row
    MsgBox "Process completed successfully with no errors and " & k & " rows."
    oLogger.Logger "Done", "RunAllCommands()", "Rate Reset has completed all commands with no errors and " & k & " rows."
    Set oLogger = Nothing
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger "Error", "RunAllCommands()", "The following error occured: " & Err.Description
    Set oLogger = Nothing
End Sub




