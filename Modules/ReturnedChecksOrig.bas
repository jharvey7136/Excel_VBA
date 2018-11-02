Attribute VB_Name = "ReturnedChecksOrig"
'
' ----------- RETURNED CHECKS -----------
'   This module does formatting along with a case replacement to change the
' returned reason code to the description. Not much heavy lifting, but it
' does save some time. This will also mask off the account number except for
' the last four digits.
'
'   DYNAMIC: built to: handle any amount of rows, calculate the correct file name based on any production date,
' calculate the correct directory to save into by calculating the correct month and year, removes unneeded rows,
' saves the file, replaces all return codes with descriptions in any order, and includes error handling
'


Sub ResetFilters()
'
' Function to clear all filter. Does not remove filters, just shows all data
'
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub


Sub dateFormat()
'
' Date Format
' Subroutine to format dates to 'Short Date' format
'
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    
End Sub

Sub numFormat()
'
' Acct Number Format
' Subroutine to format accout numbers to 'number' with no decimal places
'
    Columns("B:B").Select
    Selection.NumberFormat = "0"
End Sub

Sub moneyFormat()
'
' Currency Format
' Subroutine to format dollars with negative in red parenthesis
'
    ' Account Balance
    Columns("F:F").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
    ' Available Balance
    Columns("G:G").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
    ' Amount of Item
    Columns("H:H").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    
    ' Fee Amount
    Columns("I:I").Select
    Selection.NumberFormat = "$#,##0_);[Red]($#,##0)"
    
    ' Fee Waived
    Columns("J:J").Select
    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

End Sub

Sub RetReasonReplace()
'
' Subroutine to replace return reason codes with thier descriptions.
' Loop through the return reason cells and pass their value by reference
' to the caseSwitch function
'
    On Error GoTo eh
    
    Dim myDataRng As Range
    Dim cell As Range
    Dim retCd As String

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
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in retReasonReplace(): " & Err.Description

End Sub

Function CaseSwitch(ByRef cd As String)
'
' Case function: accepts a string as an argument by reference (returned reason code),
' Whichever code is passed in will match and return the description
'
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
              
          Case "R10"
              result = "Customer Advises Not Authorized-Item is Ineligible-Notice Not Provided-Signatures-Not Genuine or Item Altered-Approved 10-9-2001"
              
          Case "U"
              result = "Stale Dated"
          
    End Select
    
    CaseSwitch = result
    
Done:
    Exit Function
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in caseSwitch(): " & Err.Description

End Function

Sub TurnFeePositive()
'
' Subroutine to turn negative fee amounts into positive amounts
'
    Dim fees As Range
    Dim feeCell As Range

    ' Set the dynamic range of the fees
    Set fees = Range("I2:I" & Cells(Rows.Count, "I").End(xlUp).Row)
    
    ' Do one pass through range, changing negative fees to positive
    For Each feeCell In fees
        feeCell.Value = 35
    Next feeCell

End Sub

Sub MaskAccountNumbers()
'
' Subroutine to mask off all but the last four digits in account numbers with X's
'
    On Error GoTo eh
    
    Dim acctStr As String, lastFour As String, maskX As String
    Dim acctCol As Range
    Dim acctCell As Range
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
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in maskAccountNumbers(): " & Err.Description
    
End Sub

Sub SaveFile()
'
' Subroutine to save the workbook as a csv based on what month/day the post date says
'
    On Error GoTo eh
    saveFileName = GetSaveFileName()

    'save workbook as csv
    ActiveWorkbook.SaveAs fileName:= _
        saveFileName, FileFormat:=xlCSV, CreateBackup:=False
        
Done:
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured trying to saveFile(): " & Err.Description

End Sub

Function GetSaveFileName() As String
'
' Function to dynamically determine the file name string based on the prod date in column A2.
' The directory to be saved in will be in 'mmmmyy' format: October2018
' The file name date will be 'mm_dd'  eg. Returned_Checks_10_22
'
    On Error GoTo eh

    ' variables
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
    Exit Function
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in getSaveFileName(): " & Err.Description

End Function

Sub FilterAndDelete()
'
' Subroutine to test for extra, un-needed account numbers which sometimes sneak in with the query.
' If there are any, they will be removed, otherwise if there are none, exit the subroutine
'
    On Error GoTo eh
    
    ' Determine number of rows
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

    ' reset the filter to show the remaining rows
    ResetFilters
    
    ' remove autofilter
    Range("A1").Select
    Selection.AutoFilter
    
Done:
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in filterAndDelete():  " & Err.Description

End Sub

Sub RunAllReturnedChecks()
'
' One stop subroutine to run all the routines with the click of a button
'
    On Error GoTo eh

    dateFormat          ' post date = short date format
    numFormat           ' account number = number format no decimal
    moneyFormat         ' balances and fees = currency
    RetReasonReplace    ' replace return code with return description
    TurnFeePositive     ' -35 -> 35
    MaskAccountNumbers  ' mask off all but last four digits of acct number with X's
    FilterAndDelete     ' remove random lone account numbers
    
    ' Resize all columns for visual appearance
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    SaveFile            ' determine file name and save file accordingly
    
Done:
    MsgBox "Process completed successfully!" & vbNewLine & vbNewLine & "File saved to: " & GetSaveFileName()
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured in RunAllReturnedChecks(): " & Err.Description
    
End Sub

