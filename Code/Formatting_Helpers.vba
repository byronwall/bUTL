Attribute VB_Name = "Formatting_Helpers"
'''this module contains code related to formatting and other cell value stuff

''generates randoms strings of letters
Function RandLetters(count As Integer) As String

    Dim i As Integer
    
    Dim letters() As String
    ReDim letters(1 To count)
    
    For i = 1 To count
        letters(i) = Chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters, "")

End Function

'''added 2015 06 04
'''TODO some work on the Ranges to generalize
Sub ColorForUnique()

    'must add a reference to Microsoft Scripting Runtime
    Dim dict As New Scripting.Dictionary

    'build range from block of data
    'only check columns F:K for matches
    Dim rng_match As Range
    Set rng_match = Intersect( _
        Range("B2:M8"), _
        Range("F:K"))

    Dim rng_row As Range
    For Each rng_row In rng_match.rows

        Dim id As String
        id = Join(Application.Transpose(Application.Transpose(rng_row.Value)), "")

        If Not dict.Exists(id) Then
            dict.Add id, RGB(Application.RandBetween(0, 255), Application.RandBetween(0, 255), Application.RandBetween(0, 255))
        End If

        rng_row.EntireRow.Interior.Color = dict(id)
    Next rng_row
End Sub

'''this code is used to apply pretty looking number formats
Sub CreateConditionalsForFormatting()
    
    'add these in as powers of 3, starting at 1 = 10^0
    Dim arr_markers As Variant
    arr_markers = Array("", "k", "M", "B")
    
    For i = UBound(arr_markers) To 0 Step -1
        
        With Selection.FormatConditions.Add(xlCellValue, xlGreaterEqual, 10 ^ (3 * i))
            .NumberFormat = "0" & Application.WorksheetFunction.Rept(",", i) & " "" " & arr_markers(i) & """"
        End With
        
    Next

End Sub



Sub OutputColors()

    For i = 1 To 10
        ActiveCell.Offset(i).Interior.Color = Chart_GetColor(i)
    Next i

End Sub


Sub SelectedToValue()

    Dim rng As Range
    
    For Each rng In Selection
        rng.Value = rng.Value
    Next rng

End Sub

Sub Selection_ColorWithHex()

    'will color the cell with the HEX value it includes
    Dim rng_cell As Range
    
    For Each rng_cell In Selection
    
        rng_cell.Interior.Color = RGB(WorksheetFunction.Hex2Dec(Mid(rng_cell.Value, 2, 2)), WorksheetFunction.Hex2Dec(Mid(rng_cell.Value, 4, 2)), WorksheetFunction.Hex2Dec(Mid(rng_cell.Value, 6, 2)))
    
    Next rng_cell

End Sub

Sub SplitIntoColumns()

    Dim rng_input As Range
    
    Set rng_input = Intersect(Selection, ActiveSheet.UsedRange)
    
    Dim rng_cell As Range
    
    Dim str_delim As String
    str_delim = Application.InputBox("What is the delimeter?", , ",")
    
    For Each rng_cell In rng_input
    
        Dim arr_parts As Variant
        arr_parts = Split(rng_cell, str_delim)
        
        Dim var_part As Variant
        For Each var_part In arr_parts
            
            Set rng_cell = rng_cell.Offset(, 1)
            rng_cell = var_part
            
        Next var_part
    
    Next rng_cell

End Sub

Sub ExtendArrayFormulaDown()

    'Find the current array formula
    
    Dim rng_areas As Range
    Dim rng_area As Range
    Dim rng_cell As Range
    
    Application.ScreenUpdating = False
    
    Set rng_areas = Selection
    
    For Each rng_area In rng_areas.Areas
        For Each rng_cell In rng_area.Cells
        
            If rng_cell.HasArray Then
            
                Dim str_formula As String
                str_formula = rng_cell.FormulaArray
                
                Dim rng_array_start As Range
                Dim rng_array_new_end As Range
                
                Set rng_array_start = rng_cell.CurrentArray.Cells(1, 1)
                Set rng_array_new_end = rng_array_start.Offset(0, -1).End(xlDown).Offset(0, 1)
                
                rng_cell.CurrentArray.Formula = ""
                
                Range(rng_array_start, rng_array_new_end).FormulaArray = str_formula
            
            End If
        
        Next rng_cell
    Next rng_area
    
    
    'Find the range of the new array formula
    'Save current formula and clear it out
    'Apply the formula to the new range

End Sub

Public Sub Colorize()
    'this sub needs to be fully figured out.  it appears to change color when values change?
    Dim prev As Variant
    Dim cell As Range
    
    Dim active As Range
    Set active = Selection

    prev = active.Cells(1)
    
    Dim flip As Boolean

    For Each cell In Intersect(active.Columns(1).Cells, ActiveSheet.UsedRange)
        If cell <> prev Then
            flip = Not flip
            prev = cell
        End If
        
        If flip Then
            Intersect(cell.EntireRow, ActiveSheet.UsedRange).Interior.Color = RGB(200, 200, 200)
        Else
            Intersect(cell.EntireRow, ActiveSheet.UsedRange).Interior.ColorIndex = xlNone
        End If
    
    Next cell
    

End Sub

Public Sub CategoricalColoring()

    Dim rng As Range
    Dim cell As Range
    
    Dim rng_selection As Range
    Set rng_selection = Application.InputBox("Select range to color", Type:=8)
    
    Dim rng_colors As Range
    Set rng_colors = Application.InputBox("Select range with colors", Type:=8)
        
    Application.ScreenUpdating = False
    
    For Each rng In rng_selection
        Dim row As Variant
        row = Application.Match(rng, rng_colors, 0)
        If IsNumeric(row) Then
            rng.Interior.Color = rng_colors.Cells(row).Interior.Color
        End If
        
    Next rng
    
    Application.ScreenUpdating = True

End Sub

Sub TrimSelection()

    Dim cell As Range
    For Each cell In Intersect(Selection, ActiveSheet.UsedRange)
        cell.Value = Trim(cell.Value)
    Next cell

End Sub

Sub MakeHyperlinks()

    For Each cell In Selection
        ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=cell
    Next cell

End Sub

Sub ConvertToNumber()

    Dim cell As Range
    Dim sel As Range
    
    Set sel = Selection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each cell In Intersect(sel, ActiveSheet.UsedRange)
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            cell.Value = CDbl(cell.Value)
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub SplitAndKeep(delim As Variant, keep As Variant)
    
    Dim cell As Range
    Dim parts As Variant
    
    For Each cell In Intersect(Selection, ActiveSheet.UsedRange)
        parts = Split(cell, delim)
        If UBound(parts) >= keep Then
            cell = parts(keep)
        End If
    Next cell

End Sub

Sub CombineCells()

    Dim rng_input As Range, rng_output As Range
    Dim str_delim As String
    
    Dim str_output As String
    
    str_output = ""
    
    Set rng_input = Application.InputBox("Select the range of cells to combine:", Type:=8)
    str_delim = Application.InputBox("Delimeter:")
    Set rng_output = Application.InputBox("Select the output range:", Type:=8)
    
    Dim arr_values As Variant
    arr_values = Application.Transpose(Application.Transpose(rng_input.Value))
    
    rng_output = Join(arr_values, str_delim)

End Sub

Sub SplitIntoRows()

    Dim rng_out As Range
    
    Dim rng_in As Range
    Set rng_in = Selection
    
    Set rng_out = Application.InputBox("Select output corner", Type:=8)
    
    Dim part As Variant
    Dim int_row As Integer
    int_row = 0
    Dim rng_cell As Range
    
    For Each rng_cell In rng_in.SpecialCells(xlCellTypeVisible)
        Dim parts As Variant
        parts = Split(rng_cell, vbLf)
        
        For Each part In parts
            rng_out.Offset(int_row) = part
            
            int_row = int_row + 1
        Next part
    Next rng_cell
End Sub
