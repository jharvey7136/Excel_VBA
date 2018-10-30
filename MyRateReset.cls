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
Private bLog As Boolean
Private sLogFile As String

' properties
Property Let LogBool(bLogger As Boolean)
    bLog = bLogger
End Property

Property Get LogBool() As Boolean
    LogBool = bLog
End Property

Property Let LogFilePath(sPath As String)
    sLogFile = sPath
End Property

Property Get LogFilePath() As String
    LogFilePath = sLogFile
End Property

Private Sub Class_Initialize()
'    If bLog = True And sLogFile <> "" Then
'        oLogger.logFileName = sLogFile
'        oLogger.Logger "TITLE", "BEGINNING RATE RESET: ", "LOGGING TIME, TYPE, SOURCE, AND DETAILS" & vbNewLine
'    End If
End Sub

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

Public Sub RemainingLoanTerm()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "RemainingLoanTerm()", "Calculating remaining loan term..."
    End If
    UpdateProgressBar 2, 100, , "Calculating Remaining Loan Term.."
    'Application.Wait (Now + TimeValue("0:00:01"))
    
' Compute (prod date - dateOfNote) = number of days loan has been active
    Range("AC2").FormulaR1C1 = "=RC[-2]-RC[-12]"
    Range("AC2").AutoFill Destination:=Range("AC2:AC" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 4, 100, , "Calculating Remaining Loan Term.."
    'Application.Wait (Now + TimeValue("0:00:01"))
    
' Convert days to months = (days active * 0.032855) = number of months loan has been active
    Range("AD2").FormulaR1C1 = "=RC[-1]*0.032855"
    Range("AD2").AutoFill Destination:=Range("AD2:AD" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 6, 100, , "Calculating Remaining Loan Term.."
    'Application.Wait (Now + TimeValue("0:00:01"))
    
' Compute remaining term = (original term in months - months loan has been active) = remaining loan term in months
    Range("AE2").FormulaR1C1 = "=RC[-7]-RC[-1]"
    Range("AE2").AutoFill Destination:=Range("AE2:AE" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 8, 100, , "Calculating Remaining Loan Term.."
    'Application.Wait (Now + TimeValue("0:00:01"))
    
' Paste into column K: RemainingLoanTerm
    Range("AE2:AE" & getLastRow("AA")).Copy
    Range("K2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    UpdateProgressBar 10, 100, , "Calculating Remaining Loan Term.."

Done:
    If bLog = True Then
        oLogger.Logger "Done", "RemainingLoanTerm()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "RemainingLoanTerm()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured in RemainingLoanTerm(): " & Err.Description
End Sub

Sub Auth()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "Auth()", "Calculating authorization code..."
    End If
    UpdateProgressBar 12, 100, , "Generating Authorization Code.."
    
' Fetch last 4 digits of SSN
    Range("AG2").FormulaR1C1 = "=RIGHT(RC[-8],4)+0"
    Range("AG2").AutoFill Destination:=Range("AG2:AG" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 14, 100, , "Generating Authorization Code.."
    
' Concatenate year of birth and last four of SSN: 1985 + 3746 = Auth1 = 19853746
    Range("AH2").FormulaR1C1 = "=CONCATENATE(RC[-8],TEXT(RC[-1],""0000""))"
    Range("AH2").AutoFill Destination:=Range("AH2:AH" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 16, 100, , "Generating Authorization Code.."
    
' Paste into column B: Auth1
    Range("AH2:AH" & getLastRow("AA")).Copy
    Range("B2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    UpdateProgressBar 18, 100, , "Generating Authorization Code.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "Auth()", "Subrouting finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "Auth()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub FormatRR()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "FormatRR()", "Applying proper formatting..."
    End If
    UpdateProgressBar 20, 100, , "Applying Formatting.."

' LoanNumber and RemainingLoanTerm -> Number with no decimals
    Range("I:I,K:K").NumberFormat = "0"
    UpdateProgressBar 22, 100, , "Applying Formatting.."
    
' Principal Bal, Int Rate, PI Pmnt, and Orig Loan Amt -> Number with two decimal places
    Range("R:T, W:W").NumberFormat = "0.00"
    UpdateProgressBar 24, 100, , "Applying Formatting.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "FormatRR()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "FormatRR()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub JoinData()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "JoinData()", "Joining data to trans union file..."
    End If
    UpdateProgressBar 26, 100, , "Joining TransUnion Scores.."

    Dim dCells As Range, iCells As Range
    Dim RangeDataName As String, RangeIDName As String
    
    RangeDataName = "rData"
    Set dCells = Worksheets("TransUnion").Range("A1:D27000")
    ActiveWorkbook.Names.Add Name:=RangeDataName, RefersTo:=dCells
    UpdateProgressBar 28, 100, , "Joining TransUnion Scores.."
    
    RangeIDName = "rIDs"
    Set iCells = Worksheets("TransUnion").Range("A1:A27000")
    ActiveWorkbook.Names.Add Name:=RangeIDName, RefersTo:=iCells
    UpdateProgressBar 30, 100, , "Joining TransUnion Scores.."
    
    Worksheets("Data").Range("AB1").Value2 = "TScore"
    Worksheets("Data").Range("AB2").FormulaR1C1 = "=INDEX(rData,MATCH(RC[-3],rIDs,0),4)"
    Worksheets("Data").Range("AB2").AutoFill Destination:=Range("AB2:AB" & getLastRow("AA")), Type:=xlFillDefault
    UpdateProgressBar 33, 100, , "Joining TransUnion Scores.."

    Worksheets("Data").Range("AB2:AB" & getLastRow("AA")).Copy
    Worksheets("Data").Range("AB2:AB" & getLastRow("AA")).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    UpdateProgressBar 36, 100, , "Joining TransUnion Scores.."

Done:
    If bLog = True Then
        oLogger.Logger "Done", "JoinData()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "JoinData()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub SetMaxExtWithTransScores()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "SetMaxExtWithTransScores()", "Using new credit scores to re-calculate max extension..."
    End If
    UpdateProgressBar 39, 100, , "Calculating Max Extensions.."
    
    ' variables
    Dim myDataRng As Range, cell As Range, src As Integer

    ' set the range
    Set myDataRng = Range("AB2:AB" & getLastRow("AA"))
    UpdateProgressBar 42, 100, , "Calculating Max Extensions.."
    
    ' loop through range of cells, perform case switch to get max extension in months
    For Each cell In myDataRng
        If Not IsEmpty(cell) Then
            If WorksheetFunction.IsNA(cell.Value) Then
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
    
    If bLog = True Then
        oLogger.Logger "Subroutine", "SetMaxExtWithTransScores()", "Case switch finished..."
    End If
    UpdateProgressBar 46, 100, , "Calculating Max Extensions.."

    Range("AB2:AB" & getLastRow("AA")).Copy
    Range("M2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    UpdateProgressBar 50, 100, , "Calculating Max Extensions.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "SetMaxExtWithTransScores()", "Subroutine finished"
    End If
    Exit Sub
eh:
    MsgBox Err.Source & ": Error in SetMaxExtTrans(), CaseSwitch  " & Err.Description
    If bLog = True Then
        oLogger.Logger "Error", "SetMaxExtWithTransScore()", "The following error occured: " & Err.Description
    End If
End Sub

Function CaseSwitch(ByRef score As Integer)
    On Error GoTo eh

    Dim result As Integer
    
    Select Case score
        Case 650 To 1000
            result = 36
        
        Case 600 To 649
            result = 24
            
        Case 0 To 599
            result = 12
            
        Case -1
            result = -1
    End Select
    
    CaseSwitch = result
Done:
    Exit Function
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    If bLog = True Then
        oLogger.Logger "Error", "caseSwitch", "The following error occured: " & Err.Description
    End If
End Function

Sub AdjustMaxExt()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "AdjustMaxExt()", "Adjusting max extensions if necessary..."
    End If
    UpdateProgressBar 52, 100, , "Adjusting Max Extensions.."

    Range("AJ2").FormulaR1C1 = "=IF((RC[-25]+RC[-23])>84,(84-RC[-25]),RC[-23])"
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & getLastRow("AA")), Type:=xlFillDefault
    'oLogger.Logger "Subroutine", "AdjustMaxExt()", "Testing: If RemainingLoanTerm + LoanExtensionMax > 84... Done"
    UpdateProgressBar 55, 100, , "Adjusting Max Extensions.."
   
    Range("AJ2:AJ" & getLastRow("AA")).Copy
    Range("M2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'oLogger.Logger "Subroutine", "AdjustMaxExt()", "Applying new MaxExtension results... Done"
    UpdateProgressBar 57, 100, , "Adjusting Max Extensions.."
    
    Range("M:M").NumberFormat = "0"
Done:
    If bLog = True Then
        oLogger.Logger "Done", "AdjustMaxExt()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "AdjustMaxExt()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub DeDuplicate()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "DeDuplicate()", "Scanning sheet for duplicates..."
    End If
    UpdateProgressBar 61, 100, , "Deduplicating.."

    ActiveSheet.Range("$A$1:$AJ$" & getLastRow("AA")).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, _
        7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27), Header:=xlYes
    UpdateProgressBar 66, 100, , "Deduplicating.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "DeDuplicate()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "DeDuplicate()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Sub RemoveMinMax12()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "RemoveMinMax12()", "Scanning sheet for min ext and max ext both equal to 12..."
    End If
    UpdateProgressBar 69, 100, , "Removing Min/Max = 12.."
    
    ' Filter LoanExtMin by "12"
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).AutoFilter Field:=12, Criteria1:="=12"
    UpdateProgressBar 71, 100, , "Removing Min/Max = 12.."

    ' Filter LoanExtMax by "12"
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).AutoFilter Field:=13, Criteria1:="=12"

    ' Delete only visible (filtered) rows
    ActiveSheet.Range("$A$1:$AA$" & getLastRow("A")).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    UpdateProgressBar 75, 100, , "Removing Min/Max = 12.."
    
    ResetFilters
Done:
    If bLog = True Then
        oLogger.Logger "Done", "RemoveMinMax12()", "Subroutine finished"
    End If
    Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    If bLog = True Then
        oLogger.Logger "Error", "RemoveMinMax12()", "The following error occured: " & Err.Description
    End If
End Sub

Sub FilterRemainingTerm()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "FilterRemainingTerm()", "Fetching loans with >13 months remaining..."
    End If
    UpdateProgressBar 79, 100, , "Filtering Remaining Term.."

    Range("K1").AutoFilter
    
    'Filter out only loans with greater than 13 months remaining
    ActiveSheet.Range("$A$1:$AE$" & getLastRow("AA")).AutoFilter Field:=11, Criteria1:=">13", _
        Operator:=xlAnd
    UpdateProgressBar 85, 100, , "Filtering Remaining Term.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "FilterRemainingTerm()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "FilterRemainingTerm()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
    
End Sub

Sub CleanUp()
    On Error GoTo eh
    If bLog = True Then
        oLogger.Logger "Subroutine", "Cleanup()", "Removing un-needed data, moving filtered data to new sheet..."
    End If
    UpdateProgressBar 89, 100, , "Cleaning up data.."
    
    'Copy columns A through W (only columns needed)
    Range("A1:W" & getLastRow("AA")).Copy
    UpdateProgressBar 93, 100, , "Cleaning up data.."
    
    'Paste into a new sheet
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    UpdateProgressBar 98, 100, , "Cleaning up data.."
    
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    UpdateProgressBar 100, 100, , "Cleaning up data.."
Done:
    If bLog = True Then
        oLogger.Logger "Done", "Cleanup()", "Subroutine finished"
    End If
    Exit Sub
eh:
    If bLog = True Then
        oLogger.Logger "Error", "Cleanup()", "The following error occured: " & Err.Description
    End If
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub




Private Sub Class_Terminate()
    Set oLogger = Nothing
End Sub

'------------------------------ PROGRESS BAR -------------------------------------
Sub UpdateProgressBar(n As Long, m As Long, Optional DisplayText As String, Optional calcLabel As String)

    '// Omit DisplayText to display progress percentage
    On Error GoTo ERR_HANDLE

    If n >= m Then
        ProgressBar.Hide
    Else
        If ProgressBar.Visible = False Then ProgressBar.Show
        ProgressBar![Label1].Caption = calcLabel
        ProgressBar![BoxProgress].Caption = IIf(DisplayText = "", Round(((n / m) * 10000) / 100) & "%", DisplayText)
        ProgressBar![BoxProgress].Width = (n / m) * 204
        DoEvents
    End If
    Exit Sub

ERR_HANDLE:
    Err.Clear
    ProgressBar.Hide
End Sub










