Attribute VB_Name = "TESTING"
'      %                 Integer
'      &                 Long
'      !                 Single
'      #                 Double
'      $                 String
'      @                 Currency
Sub TestDir()
    Dim sPath As String
    Dim Ops As New UsefulOperations
    
    sPath = "C:\temp\dirTest"
    
    Ops.TestDirMakeDir sPath


End Sub



Sub Analyize()

    Dim oRateReset As New MyRateReset
    
    oRateReset.EndingAnalysis

End Sub



Sub CopySheet()

    
    
    Dim oUsefulOps As New UsefulOperations
    oUsefulOps.CopySheet


End Sub





Sub lereta()
'
' lereta Macro
'

'
    Columns("K:K").Select
    Selection.Copy
    Columns("C:C").Select
    ActiveSheet.Paste
    Columns("E:K").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("D:D").Select
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets.Add After:=ActiveSheet
    
    
End Sub




Sub testDate()

    Dim iMonth As Long
    Dim iDay As Long
    Dim iYear As Long
    Dim monthText As String, monthYear As String, saveDate As String

    'iMonth = Month(Date)
    'iDay = Day(Date) - 1
    
    Dim myDate As Date
    
    myDate = Range("A2").Value
    
    iMonth = Month(myDate)
    iDay = Day(myDate)
    iYear = Year(myDate)
    
    monthText = MonthName(iMonth)
    monthYear = monthText & iYear
    MsgBox (monthYear)
    
    Range("H12").Select
    Selection.Value = iMonth & "_" & iDay
    
    Range("G12").Select
    Selection.Value = monthText
    
    Range("I12").Select
    Selection.Value = iYear
    
    Range("J12").Select
    Selection.Value = monthYear
    
    'Range("F12").Select
    'Selection.Value = iMonth & "_" & iDay

End Sub

Sub AutoFill()

    LastRow = Cells(Rows.Count, "I").End(xlUp).Row
    Range("M19581").Select
    Range("M19581").AutoFill Destination:=Range("M19581:M" & LastRow), Type:=xlFillDefault


End Sub



Sub MyRateResetTest()

    Application.Calculation = xlAutomatic
    
    Dim oRateReset As New MyRateReset
    
    'oRateReset.RemainingLoanTerm
    'oRateReset.Auth
    'oRateReset.FormatRR
    'oRateReset.SetMaxExtWithTransScores
    'oRateReset.AdjustMaxExt
    'oRateReset.DeDuplicate
    'oRateReset.RemoveMinMax12
    'oRateReset.FilterRemainingTerm
    'oRateReset.Cleanup
    
    
    JoinData
    
    

End Sub


Sub JoinData()
    On Error GoTo eh

    Dim dCells As Range, iCells As Range
    Dim RangeDataName As String, RangeIDName As String
    
    RangeDataName = "rData"
    Set dCells = Worksheets("TransUnion").Range("A1:D27000")
    ActiveWorkbook.Names.Add Name:=RangeDataName, RefersTo:=dCells
    
    RangeIDName = "rIDs"
    Set iCells = Worksheets("TransUnion").Range("A1:A27000")
    ActiveWorkbook.Names.Add Name:=RangeIDName, RefersTo:=iCells
       
    
    Worksheets("Data").Range("AB1").Value2 = "TScore"
    Worksheets("Data").Range("AB2").FormulaR1C1 = "=INDEX(rData,MATCH(RC[-3],rIDs,0),4)"
    Worksheets("Data").Range("AB2").AutoFill Destination:=Range("AB2:AB" & getLastRow("AA")), Type:=xlFillDefault

Done:
    Exit Sub
eh:
    MsgBox Err.Source & ": The following error occured  " & Err.Description
End Sub

Private Function getLastRow(column As String) As Integer
    getLastRow = Cells(Rows.Count, column).End(xlUp).Row
End Function

Sub openFileExplorer()

    Dim myPath As String
    Dim folderPath As String

    folderPath = Application.ActiveWorkbook.Path
    myPath = Application.ActiveWorkbook.FullName

    Dim Msg, Style, Title, Response
    
    Msg = "Do you want to open the folder location for this workbook?" & vbNewLine & vbNewLine & _
        folderPath & vbNewLine & vbNewLine & myPath
    Style = vbYesNo + vbDefaultButton1
    Title = "Open File Explorer"
    
    Response = MsgBox(Msg, Style, Title)
    
    If Response = vbYes Then
        Shell "C:\WINDOWS\explorer.exe """ & folderPath & """", vbNormalFocus
    Else
        Exit Sub
    End If


End Sub


Sub SheetSelector()
    Const ColItems  As Long = 20
    Const LetterWidth As Long = 20
    Const HeightRowz As Long = 18
    Const SheetID As String = "__SheetSelection"
 
    Dim i%, TopPos%, iSet%, optCols%, intLetters%, optMaxChars%, optLeft%   'integers
    Dim wsDlg As DialogSheet, objOpt As OptionButton, optCaption$, objSheet As Object
        optCaption = "": i = 0
 
    'Application.ScreenUpdating = False
 
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.DialogSheets(SheetID).Delete
    Application.DisplayAlerts = True
    Err.Clear
 
    Set wsDlg = ActiveWorkbook.DialogSheets.Add
    With wsDlg
        .Name = SheetID
        .Visible = xlSheetHidden
    iSet = 0: optCols = 0: optMaxChars = 0: optLeft = 78: TopPos = 40
 
    For Each objSheet In ActiveWorkbook.Sheets
        If objSheet.Visible = xlSheetVisible Then
            i = i + 1
 
        If i Mod ColItems = 1 Then
            optCols = optCols + 1
            TopPos = 40
            optLeft = optLeft + (optMaxChars * LetterWidth)
            optMaxChars = 0
        End If
 
        intLetters = Len(objSheet.Name)
        If intLetters > optMaxChars Then optMaxChars = intLetters
            iSet = iSet + 1
            .OptionButtons.Add optLeft, TopPos, intLetters * LetterWidth, 16.5
            .OptionButtons(iSet).Text = objSheet.Name
            TopPos = TopPos + 13
 
        End If
    Next objSheet
 
    If i > 0 Then
 
        .Buttons.Left = optLeft + (optMaxChars * LetterWidth) + 24
 
        With .DialogFrame
            .Height = Application.Max(68, WorksheetFunction.Min(iSet, ColItems) * HeightRowz + 10)
            .Width = optLeft + (optMaxChars * LetterWidth) + 24
            .Caption = "Select sheet to go to"
        End With
 
    .Buttons("Button 2").BringToFront
    .Buttons("Button 3").BringToFront
    Application.ScreenUpdating = True
 
    If .Show = True Then
        For Each objOpt In wsDlg.OptionButtons
            If objOpt.Value = xlOn Then
                optCaption = objOpt.Caption
                Exit For
            End If
        Next objOpt
    End If
 
    If optCaption = "" Then
        MsgBox "You did not select a worksheet.", 48, "Cannot continue"
        Exit Sub
    Else
 
    MsgBox "You selected the sheet named ''" & optCaption & "''." & vbCrLf & "Click OK to go there.", 64, "FYI:"
    Sheets(optCaption).Activate
 
    End If
 
    End If
 
    Application.DisplayAlerts = False
        .Delete
    Application.DisplayAlerts = True
 
    End With
End Sub




Sub ProgBarTest()

    ProgressBar.Show

End Sub









Sub Macro1()
'
' power query join tables
'

'
    ActiveWorkbook.Queries.Add Name:="Table6", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Table6""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""SSN"", Int64.Type}, {""First_Name"", type text}, {""Last_Name"", type text}, {""Score_1"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table6;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Table6]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Table6_2"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("Table6_2[[#Headers],[First_Name]]").Select
    Sheets("Data").Select
    Range("B8").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$AB$132"), , xlYes).Name = _
        "Table8"
    Range("Table8[#All]").Select
    ActiveWorkbook.Queries.Add Name:="myData", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Table8""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""OfferEnds"", type datetime}, {""Auth1"", type any}, {""BorrowerFirstName"", type text}, {""BorrowerMiddleName"", type text}, {""BorrowerLastName"", type text}, {""CoBorrowerFirstname"", type text}, {""CoBorrowerMiddleName"", type text" & _
        "}, {""CoBorrowerLastName"", type text}, {""LoanNumber"", Int64.Type}, {""Email Address"", type text}, {""RemainingLoanTerm"", type any}, {""LoanExtensionMin"", Int64.Type}, {""LoanExtensionMax"", Int64.Type}, {""LoanYear"", Int64.Type}, {""LoanType"", type text}, {""Loan"", type text}, {""DateofNote"", type datetime}, {""CurrentPrincipalBalance"", type number}, {""C" & _
        "urrentIntRate"", type number}, {""CurrentPIPayment "", type number}, {""NextDueDate"", type datetime}, {""Frequency"", type text}, {""OriginalLoanAmount"", type number}, {""Original Term in Months"", Int64.Type}, {""Tax Id (Numeric)"", Int64.Type}, {""Date of Birth"", Int64.Type}, {""Prod Dt"", type datetime}, {""Score_1"", type any}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=myData;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [myData]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "myData"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("myData[[#Headers],[OfferEnds]]").Select
End Sub




