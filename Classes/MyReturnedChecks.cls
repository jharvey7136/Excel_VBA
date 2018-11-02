VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyReturnedChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ----------- RETURNED CHECKS -----------
'   This class module does formatting along with a case replacement to change the
' returned reason code to the description. Not much heavy lifting, but it
' does save some time. This will also mask off the account number except for
' the last four digits.
'
'   DYNAMIC: built to: handle any amount of rows, calculate the correct file name based on any production date,
' calculate the correct directory to save into by calculating the correct month and year, removes unneeded rows,
' saves the file, replaces all return codes with descriptions in any order, and includes error handling
'

'initialize new logger
Dim oLogger As New MyLogger
Private bLog As Boolean

Private Sub Class_Initialize()
    bLog = True
    oLogger.logFileName = "C:\temp\VB_Logger\ReturnedChecksLog.txt"
    oLogger.Logger bLog, "TITLE", "BEGINNING RETURNED CHECKS: ", "LOGGING TIME, TYPE, SOURCE, AND DETAILS" & vbNewLine
End Sub

Private Sub ResetFilters()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Public Sub Resize()
    On Error GoTo eh
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
Done:
    oLogger.Logger bLog, "Done", "Resize()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "Resize()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in Resize(): " & Err.Description
End Sub

Public Sub Formatting()
    On Error GoTo eh
    Range("A:A").NumberFormat = "m/d/yyyy"
    Range("B:B").NumberFormat = "0"
    Range("F:H, J:J").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Range("I:I").NumberFormat = "$#,##0_);[Red]($#,##0)"
Done:
    oLogger.Logger bLog, "Done", "Formatting()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "Formatting()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in Formatting(): " & Err.Description
End Sub

Public Sub RetReasonReplace()
    On Error GoTo eh
    Dim myDataRng As Range, cell As Range, retCd As String

    ' Set the dynamic range of the returned reason column
    Set myDataRng = Range("N2:N" & Cells(Rows.Count, "N").End(xlUp).Row)
    
    ' Do one pass through range, passing the value of each cell into the caseSwitch function
    For Each cell In myDataRng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) And Not cell.HasFormula Then
            retCd = cell.Value
            cell.Value = CaseSwitch(retCd)
        End If
    Next cell
Done:
    oLogger.Logger bLog, "Done", "RetReasonReplace()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "RetReasonReplace()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in RetReasonReplace(): " & Err.Description
End Sub

Private Function CaseSwitch(ByRef cd As String)
    On Error GoTo eh
    Dim result As String
    Select Case cd
          Case "A"
              result = "Insufficient Funds"
          Case "AA"
              result = "Account Frozen"
          Case "B"
              result = "Uncollected Funds"
          Case "C"
              result = "Stop Payment"
          Case "D"
              result = "Account Closed"
          Case "E"
              result = "Account Not On File"
          Case "F"
              result = "Refer To Maker"
          Case "K"
              result = "Unauthorized Signature"
          Case "M"
              result = "Lost Check"
          Case "O"
              result = "Stolen Check"
          Case "P"
              result = "Altered/Fictitious Check"
          Case "Q"
              result = "Duplicate Check"
          Case "R03"
              result = "No Account/Unable to locate account/Account Name Mismatch"
          Case "R10"
              result = "Customer Advises Not Authorized-Item is Ineligible-Notice Not Provided-Signatures-Not Genuine or Item Altered-Approved 10-9-2001"
          Case "U"
              result = "Stale Dated"
    End Select
    CaseSwitch = result
Done:
    Exit Function
eh:
    oLogger.Logger bLog, "Error", "CaseSwitch()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in CaseSwitch(): " & Err.Description
End Function

Public Sub TurnFeePositive()
    On Error GoTo eh
    Dim fees As Range, feeCell As Range

    Set fees = Range("I2:I" & Cells(Rows.Count, "I").End(xlUp).Row)
    
    ' Do one pass through range, changing negative fees to positive
    For Each feeCell In fees
        feeCell.Value = 35
    Next feeCell
Done:
    oLogger.Logger bLog, "Done", "TurnFeePositive()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "TurnFeePositive()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in TurnFeePositive(): " & Err.Description
End Sub

Public Sub MaskAccountNumbers()
    On Error GoTo eh
    
    Dim acctStr As String, lastFour As String, maskX As String
    Dim acctCol As Range, acctCell As Range
    Dim strLength As Integer, maskAmount As Integer

    ' Set the dynamic range of the acct numbers
    Set acctCol = Range("B2:B" & Cells(Rows.Count, "B").End(xlUp).Row)
    
    ' Do one pass through range
    For Each acctCell In acctCol
        ' set variables
        acctStr = acctCell.Value
        strLength = Len(acctStr)
        maskAmount = Len(acctStr) - 4
        lastFour = Right(acctStr, 4)
        
        ' set the mask string
        For i = 1 To maskAmount
            maskX = maskX & "X"
        Next i
        
        ' concatinate mask with last four
        acctStr = maskX & lastFour
        acctCell.Value = acctStr
        
        ' clear the mask string
        maskX = ""
    Next acctCell
Done:
    oLogger.Logger bLog, "Done", "MaskAccountNumbers()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "MaskAccountNumbers()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in MaskAccountNumbers(): " & Err.Description
End Sub

Public Sub SaveFile()
    On Error GoTo eh
    saveFileName = GetSaveFileName()
    
    Dim answer As Integer
    
    answer = MsgBox("Do you want to save this file?", vbYesNo + vbQuestion, "Save File")
    If answer = vbYes Then
    
        'TODO: check if directory exists, if not, create one
    
        ActiveWorkbook.SaveAs fileName:= _
        saveFileName, FileFormat:=xlCSV, CreateBackup:=False
        MsgBox "File saved to: " + saveFileName, vbOKOnly, "Save Successful"
        oLogger.Logger bLog, "FileSaved", "SaveFile()", "File successfully saved"
    Else
        oLogger.Logger bLog, "FileNotSaved", "SaveFile()", "File was not saved"
    End If
Done:
    oLogger.Logger bLog, "Done", "SaveFile()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "SaveFile()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured trying to SaveFile(): " & Err.Description
End Sub

Private Function GetSaveFileName() As String
    On Error GoTo eh

    Dim fileName As String, monthDay As String, monthYear As String
    Dim iMonth As Long, iDay As Long, iYear As Long
    Dim myDate As Date
    
    ' get date from prod date, cell A2
    myDate = Range("A2").Value
    
    ' get just the month and date integers
    iMonth = Month(myDate)
    iDay = Day(myDate)
    iYear = Year(myDate)
    
    ' concatinate with underscore
    monthDay = iMonth & "_" & iDay
    monthYear = MonthName(iMonth) & iYear
    
    ' set the file name string
    fileName = "R:\Core Systems\John Harvey\Returned Checks Daily\" & monthYear & "\Returned_Checks_" & monthDay & ".csv"
    
    GetSaveFileName = fileName
Done:
    oLogger.Logger bLog, "Done", "GetSaveFileName()", "Subroutine finished"
    Exit Function
eh:
    oLogger.Logger bLog, "Error", "GetSaveFileName()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in GetSaveFileName(): " & Err.Description
End Function

Public Sub FilterAndDelete()
    On Error GoTo eh
    preLastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Filter Post Date by 'blanks' (want to remove random solo account numbers that dont belong)
    ActiveSheet.Range("$A$1:$T$" & preLastRow).AutoFilter Field:=1, Criteria1:="="

    ' Get new number of rows (filtered)
    postLastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    ' if there are no extra random acct numbers to remove then exit
    If postLastRow = 1 Then
        ResetFilters
        Range("A1").Select
        Selection.AutoFilter
        Exit Sub
    End If
        
    ' Remove only visible rows after the filter
    ActiveSheet.Range("$A$1:$T$" & postLastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete

    ResetFilters

    Range("A1").Select
    Selection.AutoFilter
Done:
    oLogger.Logger bLog, "Done", "FilterAndDelete()", "Subroutine finished"
    Exit Sub
eh:
    oLogger.Logger bLog, "Error", "FilterAndDelete()", "The following error occured: " & Err.Description
    MsgBox Err.Source & ": The following error occured in FilterAndDelete():  " & Err.Description
End Sub

Private Sub Class_Terminate()
    Set oLogger = Nothing
End Sub
