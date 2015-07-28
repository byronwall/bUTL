Attribute VB_Name = "Formatting_Helpers"
'---------------------------------------------------------------------------------------
' Module    : Formatting_Helpers
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : contains code related to formatting and other cell value stuff
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : CategoricalColoring
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Applies the formatting from one range to another if cell value's match
'---------------------------------------------------------------------------------------
'
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


'---------------------------------------------------------------------------------------
' Procedure : ColorForUnique
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds a unique color to each unique value in a range
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub ColorForUnique()

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

'---------------------------------------------------------------------------------------
' Procedure : Colorize
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates an alternating color band based on cell values
'---------------------------------------------------------------------------------------
'
Public Sub Colorize()

    Dim rngToColor As Range
    On Error GoTo errHandler
    Set rngToColor = Application.InputBox("Select range to color", Type:=8)
    Dim lastrow As Integer
    lastrow = rngToColor.rows.count

    likevalues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)

    If likevalues = vbNo Then

        For i = 1 To lastrow
            If i Mod 2 = 0 Then
                rngToColor.rows(i).Interior.Color = RGB(200, 200, 200)
            Else: rngToColor.rows(i).Interior.ColorIndex = xlNone
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
                rngToColor.rows(i).Interior.Color = RGB(200, 200, 200)
            Else: rngToColor.rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CombineCells
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Takes a row of values and converts them to a single column
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : ConvertToNumber
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Forces all numbers stored as text to be converted to actual numbers
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : CopyTranspose
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Takes a range of cells and does a copy/tranpose
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub CopyTranspose()
'Get range to transpose
    Dim rngSrc As Range
    Set rngSrc = Application.InputBox("What range do you want to transpose?", Type:=8)
    'Obtain size of range
    Dim rRow As Integer
    rRow = rngSrc.rows.count

    Dim rCol As Integer
    rCol = rngSrc.Columns.count

    'Create array and resize to the size of the range
    Dim arrTranspose() As Variant
    ReDim arrTranspose(1 To rCol, 1 To rRow)
    For i = 1 To rRow
        For j = 1 To rCol
            arrTranspose(j, i) = rngSrc.Cells(i, j)
        Next j
    Next i

    'Where will we put it
    Dim rngDest As Range
    Set rngDest = Application.InputBox("Where would you like this to output?", Type:=8)

    'Transpose it
    rngDest.Resize(UBound(arrTranspose, 1), UBound(arrTranspose, 2)).Value = arrTranspose

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateConditionalsForFormatting
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates a set of conditional formats for order of magnitude numbers
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : ExtendArrayFormulaDown
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Takes an array formula and extends it down as far as the range on its right goes
'---------------------------------------------------------------------------------------
'
Sub ExtendArrayFormulaDown()

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

'---------------------------------------------------------------------------------------
' Procedure : MakeHyperlinks
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Converts a set of cells to hyperlink to their cell value
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : OutputColors
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Outputs the list of chart colors available
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub OutputColors()

    For i = 1 To 10
        ActiveCell.Offset(i).Interior.Color = Chart_GetColor(i)
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectedToValue
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Forces a cell to take on its value.  Removes formulas.
'---------------------------------------------------------------------------------------
'
Sub SelectedToValue()

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

'---------------------------------------------------------------------------------------
' Procedure : Selection_ColorWithHex
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Colors a cell based on the hex value stored in the cell
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub Selection_ColorWithHex()

    Dim c As Range
    Dim rngToColor As Range
    On Error GoTo errHandler
    Set rngToColor = Application.InputBox("Select the range of cells to color:", Type:=8)

    For Each c In rngToColor

        c.Interior.Color = RGB(WorksheetFunction.Hex2Dec(Mid(c.Value, 2, 2)), _
            WorksheetFunction.Hex2Dec(Mid(c.Value, 4, 2)), _
            WorksheetFunction.Hex2Dec(Mid(c.Value, 6, 2)))

    Next c
    Exit Sub
errHandler:
    MsgBox ("No selection made!")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SplitAndKeep
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Reduces a cell's value to one item returned from Split
'---------------------------------------------------------------------------------------
'
Sub SplitAndKeep(delim As Variant, vItemToKeep As Variant)
    Dim vParts As Variant
    Dim rngCell As Range

    On Error GoTo SplitAndKeep_Error

    For Each rngCell In Intersect(Selection, ActiveSheet.UsedRange)
        vParts = Split(rngCell, delim)
        If UBound(vParts) >= vItemToKeep Then
            rngCell.Value = vParts(vItemToKeep)
        End If
    Next rngCell

    On Error GoTo 0
    Exit Sub

SplitAndKeep_Error:
    MsgBox "Check that a valid Range is selected"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SplitIntoColumns
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Splits a cell into columns next to it based on a delimeter
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : SplitIntoRows
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Splits a cell with return characters into multiple rows with no returns
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : TrimSelection
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Trims whitespace from a cell's value
'---------------------------------------------------------------------------------------
'
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

