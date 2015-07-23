Attribute VB_Name = "Formatting_Helpers"
'''this module contains code related to formatting and other cell value stuff

''generates randoms strings of letters
'+changed variable name from "count"
Public Function RandLetters(num As Integer) As String

    Dim i As Integer
    
    Dim letters() As String
    ReDim letters(1 To num)
    
    For i = 1 To num
        letters(i) = Chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function
Public Sub CategoricalColoring()
    '+Get User Input
    Dim rngToColor As Range
    On Error GoTo errHandler
    Set rngToColor = Application.InputBox("Select range to color", Type:=8)
    
    Dim rngColors As Range
    Set rngColors = Application.InputBox("Select range with colors", Type:=8)
        
    '+Do Magic
    Application.ScreenUpdating = False
    Dim c As Range
    Dim varRow As Variant
    
    For Each c In rngToColor
        varRow = Application.Match(c, rngColors, 0)
        '+ Matches font style as well as interior color
        If IsNumeric(varRow) Then
            c.Font.FontStyle = rngColors.Cells(varRow).Font.FontStyle
            c.Font.Color = rngColors.Cells(varRow).Font.Color
            '+Skip interior color if there is none
            If Not rngColors.Cells(varRow).Interior.ColorIndex = xlNone Then
                c.Interior.Color = rngColors.Cells(varRow).Interior.Color
            End If
        End If
    Next c
    '+ If no fill, restore gridlines
    rngToColor.Borders.LineStyle = xlNone
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub
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
Public Sub Colorize()
   
    Dim rngToColor As Range
    On Error GoTo errHandler
    Set rngToColor = Application.InputBox("Select range to color", Type:=8)
    Dim lastrow As Integer
    lastrow = rngToColor.Cells(rows.count, "A").End(xlUp).row
    
    likevalues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)
    
    If likevalues = vbNo Then
    
        For i = 1 To lastrow
            If i Mod 2 = 0 Then
                rows(i).EntireRow.Interior.Color = RGB(200, 200, 200)
            Else: rows(i).EntireRow.Interior.ColorIndex = xlNone
            End If
        Next
    End If
    
    
    If likevalues = vbYes Then
    Dim flip As Boolean
    For i = 2 To lastrow
        If rngToColor.Cells(i, 1) <> rngToColor.Cells(i - 1, 1) Then
            flip = Not flip
        End If
        
        If flip Then
            rows(i).EntireRow.Interior.Color = RGB(200, 200, 200)
        Else: rows(i).EntireRow.Interior.ColorIndex = xlNone
        End If
    Next
    End If
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub

Sub CombineCells()

    Dim rngInput As Range
    On Error GoTo errHandler
    Set rngInput = Application.InputBox("Select the range of cells to combine:", Type:=8)
    
    Dim strDelim As String
    strDelim = Application.InputBox("Delimeter:")
    If strDelim = "" Then GoTo errHandler
    If strDelim = "False" Then GoTo errHandler
    Dim rngOutput As Range
    Set rngOutput = Application.InputBox("Select the output range:", Type:=8)
    
    'I don't understand the goal here
    Dim arr_values As Variant
    arr_values = Application.Transpose(Application.Transpose(rngInput.Value))
    
    rngOutput = Join(arr_values, strDelim)
    Exit Sub
errHandler:
    MsgBox ("No Range or Delimiter Selected!")
End Sub

Sub ConvertToNumber()
    '+I can't tell the goal of this one

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

'''this code is used to apply pretty looking number formats
Sub CreateConditionalsForFormatting()
    On Error GoTo errHandler
    Dim rngInput As Range
    Set rngInput = Application.InputBox("Select the range of cells to convert:", Type:=8)
    'add these in as powers of 3, starting at 1 = 10^0
    Dim arrMarkers As Variant
    arrMarkers = Array("", "k", "M", "B")
    
    For i = UBound(arrMarkers) To 0 Step -1
        
        With rngInput.FormatConditions.Add(xlCellValue, xlGreaterEqual, 10 ^ (3 * i))
            .NumberFormat = "0" & Application.WorksheetFunction.Rept(",", i) & " "" " & arrMarkers(i) & """"
        End With
        
    Next
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub
Sub ExtendArrayFormulaDown()

    'Find the current array formula
    '+ I don't understand this one either - it's extending an array formula down indefinitely?
    Dim rngArrForm As Range
    Dim RngArea As Range
    
    
    Application.ScreenUpdating = False
    
    Set rngArrForm = Selection
    
    For Each RngArea In rngArrForm.Areas
        For Each c In RngArea.Cells
        
            If c.HasArray Then
            
                Dim strFormula As String
                strFormula = c.FormulaArray
                
                Dim arrStart As Range
                Dim arrEnd As Range
                
                Set arrStart = c.CurrentArray.Cells(1, 1)
                Set arrEnd = arrStart.Offset(0, -1).End(xlDown).Offset(0, 1)
                
                c.CurrentArray.Formula = ""
                
                Range(arrStart, arrEnd).FormulaArray = strFormula
            
            End If
        
        Next c
    Next RngArea
    
    
    'Find the range of the new array formula
    'Save current formula and clear it out
    'Apply the formula to the new range

End Sub

Sub MakeHyperlinks()
    '+Changed to inputbox
    On Error GoTo errHandler
    Dim rngEval As Range
    Set rngEval = Application.InputBox("Select the range of cells to convert to hyperlink:", Type:=8)
    For Each c In rngEval
        ActiveSheet.Hyperlinks.Add Anchor:=c, Address:=c
    Next c
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub

Sub OutputColors()
    'function chart_getcolor() is in Chart_Helpers module
    For i = 1 To 10
        ActiveCell.Offset(i).Interior.Color = Chart_GetColor(i)
    Next i

End Sub


Sub SelectedToValue()
    'Converts calculated values to static values
    Dim rng As Range
    On Error GoTo errHandler
    Set rng = Application.InputBox("Select the formulas you'd like to convert to static values::", Type:=8)
    
    For Each c In rng
        c.Value = c.Value
    Next c
    Exit Sub
errHandler:
    MsgBox ("No selection made!")
End Sub

Sub Selection_ColorWithHex()

    'will color the cell with the HEX value it includes
    Dim c As Range
    Dim rngToColor As Range
    On Error GoTo errHandler
    Set rngToColor = Application.InputBox("Select the range of cells to color:", Type:=8)

    For Each c In rngToColor
    
        c.Interior.Color = RGB(WorksheetFunction.Hex2Dec(Mid(c.Value, 2, 2)), WorksheetFunction.Hex2Dec(Mid(c.Value, 4, 2)), WorksheetFunction.Hex2Dec(Mid(c.Value, 6, 2)))
    
    Next c
    Exit Sub
errHandler:
    MsgBox ("No selection made!")
End Sub

Sub SplitAndKeep(delim As Variant, keep As Variant)
    Dim parts As Variant
    
    For Each c In Intersect(Selection, ActiveSheet.UsedRange)
        parts = Split(c, delim)
        If UBound(parts) >= keep Then
            c = parts(keep)
        End If
    Next cell

End Sub

Sub SplitIntoColumns()

    Dim rngInput As Range
    
    Set rngInput = Intersect(Selection, ActiveSheet.UsedRange)
    
    Dim c As Range
    
    Dim strDelim As String
    strDelim = Application.InputBox("What is the delimeter?", , ",", vbOKCancel)
    If strDelim = "" Then GoTo errHandler
    If strDelim = "False" Then GoTo errHandler
    For Each c In rngInput
    
        Dim arrParts As Variant
        arrParts = Split(c, strDelim)
        
        Dim varPart As Variant
        For Each varPart In arrParts
            
            Set c = c.Offset(, 1)
            c = varPart
            
        Next varPart
    
    Next c
    Exit Sub
errHandler:
    MsgBox ("No Delimiter Defined!")
End Sub

Sub SplitIntoRows()

    Dim rngOutput As Range
    
    Dim rngInput As Range
    Set rngInput = Selection
    
    Set rngOutput = Application.InputBox("Select output corner", Type:=8)
    
    Dim varPart As Variant
    Dim iRow As Integer
    iRow = 0
    Dim c As Range
    
    For Each c In rngInput.SpecialCells(xlCellTypeVisible)
        Dim varParts As Variant
        varParts = Split(c, vbLf)
        
        For Each varPart In varParts
            rngOutput.Offset(iRow) = varPart
            
            iRow = iRow + 1
        Next varPart
    Next c
End Sub

Sub TrimSelection()
    Dim rngToTrim As Range
    On Error GoTo errHandler
    Set rngToTrim = Application.InputBox("Select the formulas you'd like to convert to static values::", Type:=8)
    
    For Each c In rngToTrim
        c.Value = Trim(c.Value)
    Next c
    Exit Sub
errHandler:
    MsgBox ("No Delimiter Defined!")
End Sub











