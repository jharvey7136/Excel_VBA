Attribute VB_Name = "FHLBI"

Sub ResetFilters()
'
' Function to clear all filter. Does not remove filters, just shows all data
'
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub



Sub FHLBI_step_one()
Attribute FHLBI_step_one.VB_Description = "Change column C to “number”. Conditional Formatting|Highlight Cell Rules|Duplicate Values|Format cells – Light Red Fill with Dark Red Text. Select Column C|Filter|Filter by color|Select “No Fill”. "
Attribute FHLBI_step_one.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FHLBI_step_one Macro
' Change column C to “number”.
' Conditional Formatting|Highlight Cell Rules|Duplicate Values|Format cells
'       – Light Red Fill with Dark Red Text.
' Select Column C|Filter|Filter by color|Select “No Fill”.
'
    ' Determine number of rows
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row

    ' Select column C -> set to "number" with no decimal places
    Columns("C:C").Select
    Selection.NumberFormat = "0"
    
    ' Set duplicate values to red fill
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    ' Filter by "No Fill"
    Range("C1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=3, Operator:= _
        xlFilterNoFill
        
End Sub
Sub FHLBI_step_two()
Attribute FHLBI_step_two.VB_Description = "Select Column D|Filter “R” and “Blanks”|Delete these rows (these are loans that were not on the last upload… or were removed from last upload) Select Column D|Filter to “U”|Change these records to “R” (these are loans that will be removed on this upload)"
Attribute FHLBI_step_two.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FHLBI_step_two Macro
' Select Column D|Filter “R” and “Blanks”|
' Delete these rows (these are loans that were not on the last upload… or were removed from last upload)
' Select Column D|Filter to “U”|Change these records to “R” (these are loans that will be removed on this upload)
'
    ' Determine number of rows (filtered)
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Filter column D by "R" and "Blanks"
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=4, Criteria1:="=R", _
        Operator:=xlOr, Criteria2:="="
    
    ' Get new number of rows (filtered)
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
    ' Remove only visible rows after the filter
    ActiveSheet.Range("$A$1:$P$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
        
    ' Re-filter column D -> only "U"
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=4, Criteria1:="=U"
    
    ' Get new number of rows (filtered)
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Replace column D on only visible rows with "R"
    Range("D2:D" & LastRow).SpecialCells(xlCellTypeVisible).Value = "R"

End Sub
Sub FHLBI_step_three()
Attribute FHLBI_step_three.VB_Description = "Select column C|Filter by color|Select Column D|Filter by “U”|Delete these rows (these were on last month’s upload sheet but had the old values…)"
Attribute FHLBI_step_three.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FHLBI_step_three Macro
' Select column C|Filter by color|Select Column D|Filter by “U”|
' Delete these rows (these were on last month’s upload sheet but had the old values…)
'
    ' clear all filters (show all data)
    ResetFilters
    
    ' Determine new number of rows
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Filter column C by color
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=3, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
        
    ' Filter column D by "U"
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=4, Criteria1:="=U"
    
    ' Delete only visible (filtered) rows
    ActiveSheet.Range("$A$1:$P$" & LastRow).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    
    
End Sub
Sub FHLBI_step_four()
Attribute FHLBI_step_four.VB_Description = "Clear all filters|Select Column D|Filter by blanks|Fill in with “U”|Clear all filters"
Attribute FHLBI_step_four.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FHLBI_step_four Macro
' Clear all filters|Select Column D|Filter by blanks|Fill in with “U”|Clear all filters
'
    ' clear all filters (show all data)
    ResetFilters
    
    ' Determine new number of rows
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Filter column D by "blanks"
    Range("D2").Select
    ActiveSheet.Range("$A$1:$P$" & LastRow).AutoFilter Field:=4, Criteria1:="="
    
    ' Replace column D on only visible rows with "U"
    Range("D2:D" & LastRow).SpecialCells(xlCellTypeVisible).Value = "U"
    
    ' Clear filters (show all data)
    ResetFilters
    
End Sub


Sub FHLBI_ExecuteAll()
' Execute all functions in sequential order

    FHLBI_step_one
    FHLBI_step_two
    FHLBI_step_three
    FHLBI_step_four

End Sub
