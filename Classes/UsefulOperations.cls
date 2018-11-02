VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsefulOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------- Test if directory exists, create it if it doesnt -----------------------

Sub TestDirMakeDir(sDirName As String)
    If Len(Dir(sDirName, vbDirectory)) = 0 Then
        MkDir sDirName
    End If
End Sub





'--------------------- Boost Macro Speed -----------------------
Sub YourSub()
    On Error GoTo eh

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Code here

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
Exit Sub
eh:
    ' Do error handling
    GoTo CleanUp
End Sub
'--------------------- Boost Macro Speed -----------------------

Sub AutoFill()
    
    On Error GoTo eh
    Dim lastRowColumn As Variant, myCol As Variant, myRow As Variant
    
    lastRowColumn = InputBox("Enter a column to determine the last row:", "Last Row")
    LastRow = Cells(Rows.Count, lastRowColumn).End(xlUp).Row

    myCol = InputBox("Enter selection cell column:")
    myRow = InputBox("Enter selection cell row:")

    Range(myCol & myRow).Select
    Range(myCol & myRow).AutoFill Destination:=Range(myCol & myRow & ":" & myCol & LastRow), Type:=xlFillDefault
Done:
    Exit Sub
eh:
    ' All errors will jump to here
    MsgBox Err.Source & ": The following error occured  " & Err.Description

End Sub


Sub DoubleSort()

    Range("A1:H23063").Select
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("G1") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom

End Sub



Sub calculateBday()
    ' Dim dobCell As Variant
    
    ' dobCell = InputBox("Enter Date of Birth cell (eg. C2):")
    Dim lastRowColumn As Variant, myCell As Variant
    
    myCell = InputBox("Enter selection cell: ")
    
    
    ActiveCell.Formula = "=INT(YEARFRAC(Range(myCell),TODAY()))"
End Sub


Sub automatic_update_on()

   Application.Calculation = xlAutomatic

End Sub

Sub automatic_update_off()

   Application.Calculation = xlManual

End Sub


Sub ResetFilters()
'
' Function to clear all filter. Does not remove filters, just shows all data
'
    On Error Resume Next
    ActiveSheet.ShowAllData

End Sub


Sub CopySheet()
    On Error GoTo eh
    
    'Dim mWB As Workbook 'master workbook
    'Dim mWS As Worksheet 'master worksheet
    'Set mWB = Application.ActiveWorkbook
    'Set mWS = mWB.ActiveSheet
    
    Cells.Select
    Cells.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    
    'mWS.Range("A1").Select
    
Done:
    Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub





