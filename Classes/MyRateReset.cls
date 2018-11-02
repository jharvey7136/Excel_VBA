VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyRateReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'                                                                                                                  '
' -------------------------------------------------- RATE RESET -------------------------------------------------- '
'   This class module does the heavy lifting in the rate reset process of preparing the                            '
' upload file. With the press of a button, all of the formatting and calculating                                   '
' is taken care of without having to do anything besides press a button.                                           '
'                                                                                                                  '
' What the class module does:                                                                                      '
'   Remaining Loan Term: calculates the amount of months left on the loan                                          '
'   Generate Auth code: building authoriziation code using SSN and birthday                                        '
'   Formatting: sets all the correct formatting to meet the upload specifications                                  '
'   JoinData: Joins the data with the trans union sheet to get the credit score                                    '
'   Set Max Extension: uses the credit score to calculate the maximum load extension in months                     '
'   Adjust Max Extension: adjusts max extension if exceeds 3 years out from current term with a max of 84 months   '
'   Deduplicate: Removes duplicates                                                                                '
'   Removes Min/Max if both equal 12: Cannot have min and max extension both equal 12 months                       '
'   Filter remaining term: only want loans with >13 months remaining                                               '
'   Clean up: pastes final version onto a new sheet                                                                '
' -----------------------------------------------------------------------------------------------------------------'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' initialize new Logger
Dim oLogger As New MyLogger
Private bLog As Boolean, bMonth As Boolean
Private sLogFile As String

'--------------------------------------- PROPERTIES ---------------------------------------
Property Let LogBool(bLogger As Boolean)    'enable/disable logging
    bLog = bLogger
End Property

Property Get LogBool() As Boolean
    LogBool = bLog
End Property
'----------------------
Property Let bFirstHalf(bFH As Boolean)     'first half or second half of month
    bMonth = bFH
End Property

Property Get bFirstHalf() As Boolean
    bFirstHalf = bMonth
End Property
'----------------------
Property Let LogFilePath(sPath As String)   'file name for logging
    sLogFile = sPath
End Property

Property Get LogFilePath() As String
    LogFilePath = sLogFile
End Property
'--------------------------------------- END PROPERTIES -----------------------------------

'--------------------------------------- HELPERS -----------------------------------
Sub SetLoggerPath()
    oLogger.logFileName = sLogFile
End Sub

Private Sub ResetFilters()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Private Function getLastRow(column As String) As Integer
    getLastRow = Cells(Rows.Count, column).End(xlUp).Row
End Function
'--------------------------------------- END HELPERS -----------------------------------

'--------------------------------------- MAIN OPERATIONS -----------------------------------
Public Sub RemainingLoanTerm()  '1. Calculate Remaining Loan Term
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "RemainingLoanTerm()", "Calculating remaining loan term..."
    
' Compute (prod date - dateOfNote) = number of days loan has been active
    Range("AC2").FormulaR1C1 = "=RC[-2]-RC[-12]"
    Range("AC2").AutoFill Destination:=Range("AC2:AC" & getLastRow("AA")), Type:=xlFillDefault
    
' Convert days to months = (days active * 0.032855) = number of months loan has been active
    Range("AD2").FormulaR1C1 = "=RC[-1]*0.032855"
    Range("AD2").AutoFill Destination:=Range("AD2:AD" & getLastRow("AA")), Type:=xlFillDefault
    
' Compute remaining term = (original term in months - months loan has been active) = remaining loan term in months
    Range("AE2").FormulaR1C1 = "=RC[-7]-RC[-1]"
    Range("AE2").AutoFill Destination:=Range("AE2:AE" & getLastRow("AA")), Type:=xlFillDefault
    
' Paste into column K: RemainingLoanTerm
    Range("AE2:AE" & getLastRow("AA")).Copy
    Range("K2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Done:
    oLogger.Logger bLog, "Done", "RemainingLoanTerm()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "RemainingLoanTerm()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in RemainingLoanTerm(): " & Err.Description
End Sub

Sub Auth() '2. Generate Authorization Code
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "Auth()", "Calculating authorization code..."
    
' Fetch last 4 digits of SSN
    Range("AG2").FormulaR1C1 = "=RIGHT(RC[-8],4)+0"
    Range("AG2").AutoFill Destination:=Range("AG2:AG" & getLastRow("AA")), Type:=xlFillDefault
    
' Concatenate year of birth and last four of SSN: 1985 + 3746 = Auth1 = 19853746
    Range("AH2").FormulaR1C1 = "=CONCATENATE(RC[-8],TEXT(RC[-1],""0000""))"
    Range("AH2").AutoFill Destination:=Range("AH2:AH" & getLastRow("AA")), Type:=xlFillDefault
    
' Paste into column B: Auth1
    Range("AH2:AH" & getLastRow("AA")).Copy
    Range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Done:
    oLogger.Logger bLog, "Done", "Auth()", "Subrouting finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "Auth()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub FormatRR() '3. Apply Formatting
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "FormatRR()", "Applying proper formatting..."

    Range("I:I,K:K").NumberFormat = "0"  'LoanNumber and RemainingLoanTerm -> Number with no decimals
    
    Range("A:A,Q:Q,U:U").NumberFormat = "m/d/yyyy"  'set short date format for dates

    Range("R:T, W:W").NumberFormat = "0.00" 'Principal Bal, Int Rate, PI Pmnt, and Orig Loan Amt -> Number with two decimal places
Done:
    oLogger.Logger bLog, "Done", "FormatRR()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "FormatRR()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub JoinData(sTrans As String, sData As String) '4. Join table with TransUnion to get credit score       'TODO: add sTransUnion, sDataSheet
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "JoinData()", "Joining data to trans union file..."

    Dim dCells As Range, iCells As Range    'range variables
    Dim RangeDataName As String, RangeIDName As String  'strings for range names
    
    RangeDataName = "rData"
    Set dCells = Worksheets(sTrans).Range("A1:D27000")    'set range as the full trans union table
    ActiveWorkbook.Names.Add Name:=RangeDataName, RefersTo:=dCells
    
    RangeIDName = "rIDs"
    Set iCells = Worksheets(sTrans).Range("A1:A27000")    'set range as just the SSNs
    ActiveWorkbook.Names.Add Name:=RangeIDName, RefersTo:=iCells
    
    Worksheets(sData).Range("AB1").Value2 = "TScore"
    Worksheets(sData).Range("AB2").FormulaR1C1 = "=INDEX(rData,MATCH(RC[-3],rIDs,0),4)"    'index and match to join based on SSNs
    Worksheets(sData).Range("AB2").AutoFill Destination:=Range("AB2:AB" & getLastRow("AA")), Type:=xlFillDefault

    Worksheets(sData).Range("AB2:AB" & getLastRow("AA")).Copy  'copy/paste values
    Worksheets(sData).Range("AB2:AB" & getLastRow("AA")).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Done:
    oLogger.Logger bLog, "Done", "JoinData()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "JoinData()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub SetMaxExtWithTransScores() '5. Calculate Max Extension based on credit score from trans union
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "SetMaxExtWithTransScores()", "Using new credit scores to re-calculate max extension..."
    
    Dim myDataRng As Range, cell As Range, src As Integer   'variables
    Set myDataRng = Range("AB2:AB" & getLastRow("AA"))  'set the range
    
    For Each cell In myDataRng          'loop through range of cells, perform case switch to get max extension in months
        If Not IsEmpty(cell) Then
            If WorksheetFunction.IsNA(cell.Value) Then  'handle 'NA' cells
                cell.Value = -1
            End If
            src = cell.Value
            cell.Value = CaseSwitch(src)
        End If
    Next cell
    
    For Each cell In myDataRng
        If cell.Value = -1 Then
            cell.Value = ""
        End If
    Next cell
    
    oLogger.Logger bLog, "Subroutine", "SetMaxExtWithTransScores()", "Case switch finished..."

    Range("AB2:AB" & getLastRow("AA")).Copy
    Range("M2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
Done:
    oLogger.Logger bLog, "Done", "SetMaxExtWithTransScores()", "Subroutine finished"
    Exit Sub
eh:
    MsgBox Err.Source & ": Error in SetMaxExtTrans(), CaseSwitch  " & Err.Description
    oLogger.Logger bLog, "Error", "SetMaxExtWithTransScore()", "The following error occured: " & Err.Description
End Sub

Function CaseSwitch(ByRef score As Integer) 'Helper function to set the max ext
    On Error GoTo eh
    Dim result As Integer
    Select Case score
        Case 650 To 1000
            result = 36
        
        Case 600 To 649
            result = 24
            
        Case 0 To 599
            result = 12
            
        Case -1         'set 'NA' to -1
            result = -1
    End Select
    
    CaseSwitch = result
Done:
    Exit Function
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger bLog, "Error", "caseSwitch", "The following error occured: " & Err.Description
End Function

Sub AdjustMaxExt() '6. Adjust the max extension if needed
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "AdjustMaxExt()", "Adjusting max extensions if necessary..."

    Range("AJ2").FormulaR1C1 = "=IF((RC[-25]+RC[-23])>84,(84-RC[-25]),RC[-23])"
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & getLastRow("AA")), Type:=xlFillDefault
    oLogger.Logger bLog, "Subroutine", "AdjustMaxExt()", "Testing: If RemainingLoanTerm + LoanExtensionMax > 84... Done"
   
    Range("AJ2:AJ" & getLastRow("AA")).Copy
    Range("M2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    oLogger.Logger bLog, "Subroutine", "AdjustMaxExt()", "Applying new MaxExtension results... Done"
    
    Range("M:M").NumberFormat = "0"
Done:
    oLogger.Logger bLog, "Done", "AdjustMaxExt()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "AdjustMaxExt()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub DeDuplicate()   '7. Remove duplicates, some loans get added twice if there are multiple credit scores on file for a person
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "DeDuplicate()", "Scanning sheet for duplicates..."

    ActiveSheet.Range("$A$1:$AJ$" & getLastRow("AA")).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, _
        7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27), Header:=xlYes
Done:
    oLogger.Logger bLog, "Done", "DeDuplicate()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "DeDuplicate()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub RemoveMinMax12() '8. Remove any rows where both the min extension and the max equal 12 months
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "RemoveMinMax12()", "Scanning sheet for min ext and max ext both equal to 12..."
    
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).AutoFilter Field:=12, Criteria1:="=12" 'Filter LoanExtMin by "12"
    
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).AutoFilter Field:=13, Criteria1:="=12" 'Filter LoanExtMax by "12"

    ' Delete only visible (filtered) rows
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    
    ResetFilters    'show all data
Done:
    oLogger.Logger bLog, "Done", "RemoveMinMax12()", "Subroutine finished"
    Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    oLogger.Logger bLog, "Error", "RemoveMinMax12()", "The following error occured: " & Err.Description
End Sub

Sub FilterRemainingTerm()   '9. Filter rows to only show loans with more than 13 months remaining
    On Error GoTo eh
    oLogger.Logger bLog, "Subroutine", "FilterRemainingTerm()", "Fetching loans with >13 months remaining..."

    Range("K1").AutoFilter  'enable filter
    
    'Filter out only loans with greater than 13 months remaining
    ActiveSheet.Range("$A$1:$AE$" & getLastRow("AA")).AutoFilter Field:=11, Criteria1:=">13", _
        Operator:=xlAnd
Done:
    oLogger.Logger bLog, "Done", "FilterRemainingTerm()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "FilterRemainingTerm()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub CleanUp()   '10. Cleanup: copy and paste filtered data to a new sheet and autosize columns
    On Error GoTo eh
    
    oLogger.Logger bLog, "Subroutine", "Cleanup()", "Removing un-needed data, moving filtered data to new sheet..."
    
    Range("A1:W" & getLastRow("AA")).Copy       'Copy columns A through W (only columns needed)
    
    Sheets.Add After:=ActiveSheet   'Paste into a new sheet
    ActiveSheet.Paste
    
    Cells.EntireColumn.AutoFit  'autofit columns
    Range("A1").Select
    ActiveSheet.Name = SetSheetName
    
    'EndingAnalysis
    
Done:
    oLogger.Logger bLog, "Done", "Cleanup()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "Cleanup()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub
'--------------------------------------- END MAIN OPERATIONS -----------------------------------

Function SetSheetName() As String
    On Error GoTo eh
    Dim sMonth As String, sHalf As String
    sMonth = Left(RateResetForm.sMonthName, 3)
    
    If bFirstHalf Then
        sHalf = "1_15"
    Else
        sHalf = "16_End"
    End If
    SetSheetName = "RateReset_" & sMonth & "18_" & sHalf
Done:
    Exit Function
eh:
    MsgBox Err.Source & ": The following error occured in MyRateReset.SetSheetName() " & Err.Description

End Function

'--------------------------------------- CLASS ROUTINES-----------------------------------
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set oLogger = Nothing
End Sub
'--------------------------------------- END CLASS ROUTINES-----------------------------------



'------------------------- DATA ANALYSIS -------------------------------
Sub EndingAnalysis()
    Dim iNumRows As Integer
    Dim siAvgInt As Single, siAvgRmTerm As Single, siAvgPI As Single, siAvgPrinBal As Single
    Dim siAvgMaxExt As Single
    
    iNumRows = (ActiveSheet.Range("A1048576").End(xlUp).Row) - 1      'number of rows
    siAvgInt = WorksheetFunction.Average(Range("S2:S" & getLastRow("S")))
    siAvgRmTerm = WorksheetFunction.Average(Range("K2:K" & getLastRow("K")))
    siAvgPI = WorksheetFunction.Average(Range("T2:T" & getLastRow("T")))
    siAvgPrinBal = WorksheetFunction.Average(Range("R2:R" & getLastRow("R")))
    siAvgMaxExt = WorksheetFunction.Average(Range("M2:M" & getLastRow("M")))
    
    
    MsgBox "Number of Rows: " & iNumRows & vbNewLine & _
           "Avg Interest Rate: " & Format((siAvgInt / 100), "Percent") & vbNewLine & _
           "Avg Remaining Term: " & Format(siAvgRmTerm, "#,##0") & " months" & vbNewLine & _
           "Avg PI Payment: " & Format(siAvgPI, "Currency") & vbNewLine & _
           "Avg Principal Balance: " & Format(siAvgPrinBal, "Currency") & vbNewLine & _
           "Avg Max Extension: " & Format(siAvgMaxExt, "#,##0") & " months"

End Sub










