Attribute VB_Name = "Formatting_Helpers"
Option Explicit

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
    Dim rangeToColor As Range
    On Error GoTo errHandler
    Set rangeToColor = GetInputOrSelection("Select Range to Color")

    Dim rangeWithColors As Range
    Set rangeWithColors = GetInputOrSelection("Select Range with Colors")

    '+Do Magic
    Application.ScreenUpdating = False
    Dim myRange As Range
    Dim variableRow As Variant

    For Each myRange In rangeToColor
        variableRow = Application.Match(myRange, rangeWithColors, 0)
        '+ Matches font style as well as interior color
        If IsNumeric(variableRow) Then
            myRange.Font.FontStyle = rangeWithColors.Cells(variableRow).Font.FontStyle
            myRange.Font.Color = rangeWithColors.Cells(variableRow).Font.Color
            '+Skip interior color if there is none
            If Not rangeWithColors.Cells(variableRow).Interior.ColorIndex = xlNone Then
                myRange.Interior.Color = rangeWithColors.Cells(variableRow).Interior.Color
            End If
        End If
    Next myRange
    '+ If no fill, restore gridlines
    rangeToColor.Borders.LineStyle = xlNone
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ColorForUnique
' Author    : @byronwall, @RaymondWise
' Date      : 2015 07 29
' Purpose   : Adds the same unique color to each unique value in a range
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub ColorForUnique()

    Dim dictKeysAndColors As New Scripting.Dictionary
    Dim dictColorsOnly As New Scripting.Dictionary
    
    Dim rangeToColor As Range

    On Error GoTo ColorForUnique_Error

    Set rangeToColor = GetInputOrSelection("Select column to color")
    Set rangeToColor = Intersect(rangeToColor, rangeToColor.Parent.UsedRange)

    'We can colorize the sorting column, or the entire row
    Dim colorEntireRow As VbMsgBoxResult
    colorEntireRow = MsgBox("Do you want to color the entire row?", vbYesNo)

    Application.ScreenUpdating = False

    Dim rngRowToColor As Range
    For Each rngRowToColor In rangeToColor.Rows

        'allow for a multi column key if intial range is multi-column
        'TODO: consider making this another prompt... might (?) want to color multi range based on single column key
        Dim identification As String
        If rngRowToColor.Columns.count > 1 Then
            identification = Join(Application.Transpose(Application.Transpose(rngRowToColor.Value)), "||")
        Else
            identification = rngRowToColor.Value
        End If

        'new value, need a color
        If Not dictKeysAndColors.Exists(identification) Then
            Dim newRGB As Long
createNewColor:
            newRGB = RGB(Application.RandBetween(50, 255), _
                             Application.RandBetween(50, 255), Application.RandBetween(50, 255))
            If dictColorsOnly.Exists(newRGB) Then
                'ensure unique colors only
                GoTo createNewColor
            End If
                
            dictKeysAndColors.Add identification, newRGB
        End If

        If colorEntireRow = vbYes Then
            rngRowToColor.EntireRow.Interior.Color = dictKeysAndColors(identification)
        Else
            rngRowToColor.Interior.Color = dictKeysAndColors(identification)
        End If
    Next rngRowToColor

    Application.ScreenUpdating = True

    On Error GoTo 0
    Exit Sub

ColorForUnique_Error:
    MsgBox "Select a valid range or fewer than 65650 unique entries."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Colorize
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates an alternating color band based on cell values
'---------------------------------------------------------------------------------------
'
Public Sub Colorize()

    Dim rangeToColor As Range
    On Error GoTo errHandler
    Set rangeToColor = GetInputOrSelection("Select range to color")
    Dim lastRow As Long
    lastRow = rangeToColor.Rows.count
    
    Dim likeValues As VbMsgBoxResult
    likeValues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)

    If likeValues = vbNo Then
        
        Dim i As Long
        For i = 1 To lastRow
            If i Mod 2 = 0 Then
                rangeToColor.Rows(i).Interior.Color = RGB(200, 200, 200)
            Else: rangeToColor.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If


    If likeValues = vbYes Then
        Dim flipColors As Boolean
        For i = 2 To lastRow
            If rangeToColor.Cells(i, 1) <> rangeToColor.Cells(i - 1, 1) Then
                flipColors = Not flipColors
            End If

            If flipColors Then
                rangeToColor.Rows(i).Interior.Color = RGB(200, 200, 200)
            Else: rangeToColor.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If
    Exit Sub
errHandler:
    MsgBox ("No Range Selected!")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CombineCells
' Author    : @byronwall, @RaymondWise
' Date      : 2015 07 24
' Purpose   : Takes a row of values and converts them to a single column
'---------------------------------------------------------------------------------------
'
Sub CombineCells()
    'collect all user data up front
    Dim rngInput As Range
    On Error GoTo errHandler
    Set rngInput = GetInputOrSelection("Select the range of cells to combine")

    Dim strDelim As String
    strDelim = Application.InputBox("Delimeter:")
    If strDelim = "" Then GoTo errHandler
    If strDelim = "False" Then GoTo errHandler
    Dim rngOutput As Range
    Set rngOutput = GetInputOrSelection("Select the output range")
    
    'Check the size of input and adjust output
    Dim y As Long
    y = rngInput.Columns.count
    
    Dim x As Long
    x = rngInput.Rows.count
    
    rngOutput = rngOutput.Resize(x, 1)
    
    'Read input rows into a single string
    Dim strOutput As String
    Dim i As Long
    For i = 1 To x
        strOutput = vbNullString
        Dim j As Long
        For j = 1 To y
            strOutput = strOutput & strDelim & rngInput(i, j)
        Next
        'Get rid of the first character (strDelim)
        strOutput = Right(strOutput, Len(strOutput) - 1)
        'Print it!
        rngOutput(i, 1) = strOutput
    Next
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
' Author    : @byronwall, @RaymondWise
' Date      : 2015 07 31
' Purpose   : Takes a range of cells and does a copy/tranpose
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub CopyTranspose()

    'If user cancels a range input, we need to handle it when it occurs
    On Error GoTo errCancel
    Dim rngSelect As Range
    
    Set rngSelect = GetInputOrSelection("Select your range")

    Dim rngOut As Range
    Set rngOut = GetInputOrSelection("Select the output corner")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim rCorner As Range
    Set rCorner = rngSelect.Cells(1, 1)

    Dim iCRow As Long
    iCRow = rCorner.row
    Dim iCCol As Long
    iCCol = rCorner.Column

    Dim iORow As Long
    Dim iOCol As Long
    iORow = rngOut.row
    iOCol = rngOut.Column

    Dim c As Range
    
    'We check for the intersection to ensure we don't overwrite any of the original data
    For Each c In rngSelect
        If Not Intersect(rngSelect, Cells(iORow + c.Column - iCCol, iOCol + c.row - iCRow)) Is Nothing Then
            MsgBox ("Your destination intersects with your data")
            Exit Sub
        End If
    Next c

    For Each c In rngSelect
        ActiveSheet.Cells(iORow + c.Column - iCCol, iOCol + c.row - iCRow).Formula = c.Formula
    Next c

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    
errCancel:
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
    Set rngInput = GetInputOrSelection("Select the range of cells to convert")
    'add these in as powers of 3, starting at 1 = 10^0
    Dim arrMarkers As Variant
    arrMarkers = Array("", "k", "M", "B")
    
    Dim i As Long
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
    
        Dim c As Range
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
    Set rngEval = GetInputOrSelection("Select the range of cells to convert to hyperlink")
    
    'TODO: choose a better variable name
    Dim c As Range
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
    
    Dim i As Long
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
    Set rng = GetInputOrSelection("Select the formulas you'd like to convert to static values")

    Dim c As Range
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
    Dim rangeToColor As Range
    On Error GoTo errHandler
    Set rangeToColor = GetInputOrSelection("Select the range of cells to color")

    For Each c In rangeToColor

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
' Date      : 2015 08 12
' Purpose   : Reduces a cell's value to one item returned from Split
'---------------------------------------------------------------------------------------
'
Sub SplitAndKeep()

    On Error GoTo SplitAndKeep_Error

    Dim rngToSplit As Range
    Set rngToSplit = GetInputOrSelection("Select range to split")
    
    If rngToSplit Is Nothing Then
        Exit Sub
    End If

    Dim delim As Variant
    delim = InputBox("What delimeter to split on?")
    
    If StrPtr(delim) = 0 Then
        Exit Sub
    End If

    Dim vItemToKeep As Variant
    vItemToKeep = InputBox("Which item to keep? (This is 0-indexed)")
    
    If StrPtr(vItemToKeep) = 0 Then
        Exit Sub
    End If

    Dim rngCell As Range
    For Each rngCell In Intersect(rngToSplit, rngToSplit.Parent.UsedRange)
        
        Dim vParts As Variant
        vParts = Split(rngCell, delim)
        
        If UBound(vParts) >= vItemToKeep Then
            rngCell.Value = vParts(vItemToKeep)
        End If
        
    Next rngCell

    On Error GoTo 0
    Exit Sub

SplitAndKeep_Error:
    MsgBox "Check that a valid Range is selected and that a number was entered for which item to keep."
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SplitIntoColumns
' Author    : @byronwall, @RaymondWise
' Date      : 2015 07 24
' Purpose   : Splits a cell into columns next to it based on a delimeter
'---------------------------------------------------------------------------------------
'
Sub SplitIntoColumns()

    Dim rngInput As Range

    Set rngInput = GetInputOrSelection("Select the range of cells to split")

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

    Set rngOutput = GetInputOrSelection("Select the output corner")

    Dim varPart As Variant
    Dim iRow As Long
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
    Set rngToTrim = GetInputOrSelection("Select the formulas you'd like to convert to static values")
    
    Dim c As Range
    For Each c In rngToTrim
        c.Value = Trim(c.Value)
    Next c
    Exit Sub
errHandler:
    MsgBox ("No Delimiter Defined!")
End Sub

