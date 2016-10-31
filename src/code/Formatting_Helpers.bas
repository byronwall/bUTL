Attribute VB_Name = "Formatting_Helpers"
Option Explicit


Public Sub CategoricalColoring()
    '---------------------------------------------------------------------------------------
    ' Procedure : CategoricalColoring
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies the formatting from one range to another if targetCell value's match
    '---------------------------------------------------------------------------------------
    '

    '+Get User Input
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select Range to Color")

    Dim coloredRange As Range
    Set coloredRange = GetInputOrSelection("Select Range with Colors")

    '+Do Magic
    Application.ScreenUpdating = False
    Dim targetCell As Range
    Dim foundRange As Variant

    For Each targetCell In targetRange
        foundRange = Application.Match(targetCell, coloredRange, 0)
        '+ Matches font style as well as interior color
        If IsNumeric(foundRange) Then
            targetCell.Font.FontStyle = coloredRange.Cells(foundRange).Font.FontStyle
            targetCell.Font.Color = coloredRange.Cells(foundRange).Font.Color
            '+Skip interior color if there is none
            If Not coloredRange.Cells(foundRange).Interior.ColorIndex = xlNone Then
                targetCell.Interior.Color = coloredRange.Cells(foundRange).Interior.Color
            End If
        End If
    Next targetCell
    '+ If no fill, restore gridlines
    targetRange.Borders.LineStyle = xlNone
    Application.ScreenUpdating = True
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
    Application.ScreenUpdating = True
End Sub



Public Sub ColorForUnique()
    '---------------------------------------------------------------------------------------
    ' Procedure : ColorForUnique
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 29
    ' Purpose   : Adds the same unique color to each unique value in a range
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim dictKeysAndColors As New Scripting.Dictionary
    Dim dictColorsOnly As New Scripting.Dictionary
    
    Dim targetRange As Range

    On Error GoTo ColorForUnique_Error

    Set targetRange = GetInputOrSelection("Select column to color")
    Set targetRange = Intersect(targetRange, targetRange.Parent.UsedRange)

    'We can colorize the sorting column, or the entire row
    Dim shouldColorEntireRow As VbMsgBoxResult
    shouldColorEntireRow = MsgBox("Do you want to color the entire row?", vbYesNo)

    Application.ScreenUpdating = False

    Dim rowToColor As Range
    For Each rowToColor In targetRange.Rows

        'allow for a multi column key if intial range is multi-column
        'TODO: consider making this another prompt... might (?) want to color multi range based on single column key
        Dim keyString As String
        If rowToColor.Columns.count > 1 Then
            keyString = Join(Application.Transpose(Application.Transpose(rowToColor.Value)), "||")
        Else
            keyString = rowToColor.Value
        End If

        'new value, need a color
        If Not dictKeysAndColors.Exists(keyString) Then
            Dim randomColor As Long
createNewColor:
            randomColor = RGB(Application.RandBetween(50, 255), _
                            Application.RandBetween(50, 255), Application.RandBetween(50, 255))
            If dictColorsOnly.Exists(randomColor) Then
                'ensure unique colors only
                GoTo createNewColor 'This is a sub-optimal way of performing this error check and loop
            End If
                
            dictKeysAndColors.Add keyString, randomColor
        End If

        If shouldColorEntireRow = vbYes Then
            rowToColor.EntireRow.Interior.Color = dictKeysAndColors(keyString)
        Else
            rowToColor.Interior.Color = dictKeysAndColors(keyString)
        End If
    Next rowToColor

    Application.ScreenUpdating = True

    On Error GoTo 0
    Exit Sub

ColorForUnique_Error:
    MsgBox "Select a valid range or fewer than 65650 unique entries."

End Sub


Public Sub Colorize()
    '---------------------------------------------------------------------------------------
    ' Procedure : Colorize
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates an alternating color band based on targetCell values
    '---------------------------------------------------------------------------------------
    '
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select range to color")
    Dim lastRow As Long
    lastRow = targetRange.Rows.count
    Dim interiorColor As Long
    interiorColor = RGB(200, 200, 200)
    
    Dim sameColorForLikeValues As VbMsgBoxResult
    sameColorForLikeValues = MsgBox("Do you want to keep duplicate values the same color?", vbYesNo)

    If sameColorForLikeValues = vbNo Then
        
        Dim i As Long
        For i = 1 To lastRow
            If i Mod 2 = 0 Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If


    If sameColorForLikeValues = vbYes Then
        Dim flipFlag As Boolean
        For i = 2 To lastRow
            If targetRange.Cells(i, 1) <> targetRange.Cells(i - 1, 1) Then flipFlag = Not flipFlag
            If flipFlag Then
                targetRange.Rows(i).Interior.Color = interiorColor
            Else: targetRange.Rows(i).Interior.ColorIndex = xlNone
            End If
        Next
    End If
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub


Public Sub CombineCells()
    '---------------------------------------------------------------------------------------
    ' Procedure : CombineCells
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 24
    ' Purpose   : Takes a row of values and converts them to a single column
    '---------------------------------------------------------------------------------------
    '
    'collect all user data up front
    Dim inputRange As Range
    On Error GoTo errHandler
    Set inputRange = GetInputOrSelection("Select the range of cells to combine")

    Dim delimiter As String
    delimiter = Application.InputBox("Delimeter:")
    If delimiter = "" Or delimiter = "False" Then GoTo delimiterError

    Dim outputRange As Range
    Set outputRange = GetInputOrSelection("Select the output range")
    
    'Check the size of input and adjust output
    Dim numberOfColumns As Long
    numberOfColumns = inputRange.Columns.count
    
    Dim numberOfRows As Long
    numberOfRows = inputRange.Rows.count
    
    outputRange = outputRange.Resize(numberOfRows, 1)
    
    'Read input rows into a single string
    Dim outputString As String
    Dim i As Long
    For i = 1 To numberOfRows
        outputString = vbNullString
        Dim j As Long
        For j = 1 To numberOfColumns
            outputString = outputString & delimiter & inputRange(i, j)
        Next
        'Get rid of the first character (delimiter)
        outputString = Right(outputString, Len(outputString) - 1)
        'Print it!
        outputRange(i, 1) = outputString
    Next
    Exit Sub
delimiterError:
    MsgBox "No Delmiter Selected!"
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub


Public Sub ConvertToNumber()
    '---------------------------------------------------------------------------------------
    ' Procedure : ConvertToNumber
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces all numbers stored as text to be converted to actual numbers
    '---------------------------------------------------------------------------------------
    '
    Dim targetCell As Range
    Dim targetSelection As Range

    Set targetSelection = Selection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each targetCell In Intersect(targetSelection, ActiveSheet.UsedRange)
        If Not IsEmpty(targetCell.Value) And IsNumeric(targetCell.Value) Then targetCell.Value = CDbl(targetCell.Value)
    Next targetCell

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Public Sub CopyTranspose()
    '---------------------------------------------------------------------------------------
    ' Procedure : CopyTranspose
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 31
    ' Purpose   : Takes a range of cells and does a copy/tranpose
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    'If user cancels a range input, we need to handle it when it occurs
    On Error GoTo errCancel
    Dim selectedRange As Range
    
    Set selectedRange = GetInputOrSelection("Select your range")

    Dim outputRange As Range
    'Need to handle the error of selecting more than one cell
    Set outputRange = GetInputOrSelection("Select the output corner")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim startingCornerCell As Range
    Set startingCornerCell = selectedRange.Cells(1, 1)

    Dim startingCellRow As Long
    startingCellRow = startingCornerCell.Row
    Dim startingCellColumn As Long
    startingCellColumn = startingCornerCell.Column

    Dim outputRow As Long
    Dim outputColumn As Long
    outputRow = outputRange.Row
    outputColumn = outputRange.Column

    Dim targetCell As Range
    
    'We check for the intersection to ensure we don't overwrite any of the original data
    'There's probably a better way to do this than For Each
    For Each targetCell In selectedRange
        If Not Intersect(selectedRange, Cells(outputRow + targetCell.Column - startingCellColumn, outputColumn + targetCell.Row - startingCellRow)) Is Nothing Then
            MsgBox "Your destination intersects with your data"
            Exit Sub
        End If
    Next targetCell

    For Each targetCell In selectedRange
        ActiveSheet.Cells(outputRow + targetCell.Column - startingCellColumn, outputColumn + targetCell.Row - startingCellRow).Formula = targetCell.Formula
    Next targetCell

errCancel:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
End Sub




Public Sub CreateConditionalsForFormatting()
    '---------------------------------------------------------------------------------------
    ' Procedure : CreateConditionalsForFormatting
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a set of conditional formats for order of magnitude numbers
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo errHandler
    Dim inputRange As Range
    Set inputRange = GetInputOrSelection("Select the range of cells to convert")
    'add these in as powers of 3, starting at 1 = 10^0
    Const ARRAY_MARKERS As String = " ,k,M,B,T,Q"
    Dim arrMarkers As Variant
    arrMarkers = Split(ARRAY_MARKERS, ",")
    
    Dim i As Long
    For i = UBound(arrMarkers) To 0 Step -1

        With inputRange.FormatConditions.Add(xlCellValue, xlGreaterEqual, 10 ^ (3 * i))
            .NumberFormat = "0.0" & Application.WorksheetFunction.Rept(",", i) & " "" " & arrMarkers(i) & """"
        End With

    Next
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub


Public Sub ExtendArrayFormulaDown()
    '---------------------------------------------------------------------------------------
    ' Procedure : ExtendArrayFormulaDown
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Takes an array formula and extends it down as far as the range on its right goes
    '---------------------------------------------------------------------------------------
    '
    Dim startingRange As Range
    Dim targetArea As Range


    Application.ScreenUpdating = False

    Set startingRange = Selection

    For Each targetArea In startingRange.Areas
    
        Dim targetCell As Range
        For Each targetCell In targetArea.Cells

            If targetCell.HasArray Then

                Dim formulaString As String
                formulaString = targetCell.FormulaArray

                Dim startOfArray As Range
                Dim endOfArray As Range

                Set startOfArray = targetCell.CurrentArray.Cells(1, 1)
                Set endOfArray = startOfArray.Offset(0, -1).End(xlDown).Offset(0, 1)

                targetCell.CurrentArray.Formula = vbNullString

                Range(startOfArray, endOfArray).FormulaArray = formulaString

            End If

        Next targetCell
    Next targetArea


    'Find the range of the new array formula
    'Save current formula and clear it out
    'Apply the formula to the new range
    Application.ScreenUpdating = True
End Sub


Public Sub MakeHyperlinks()
    '---------------------------------------------------------------------------------------
    ' Procedure : MakeHyperlinks
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Converts a set of cells to hyperlink to their targetCell value
    '---------------------------------------------------------------------------------------
    '
    '+Changed to inputbox
    On Error GoTo errHandler
    Dim targetRange As Range
    Set targetRange = GetInputOrSelection("Select the range of cells to convert to hyperlink")
    
    'TODO: choose a better variable name
    Dim targetCell As Range
    For Each targetCell In targetRange
        ActiveSheet.Hyperlinks.Add Anchor:=targetCell, Address:=targetCell
    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No Range Selected!"
End Sub


Public Sub OutputColors()
    '---------------------------------------------------------------------------------------
    ' Procedure : OutputColors
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Outputs the list of chart colors available
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Const MINIMUM_INTEGER As Long = 1
    Const MAXIMUM_INTEGER As Long = 10
    Dim i As Long
    For i = MINIMUM_INTEGER To MAXIMUM_INTEGER
        ActiveCell.Offset(i).Interior.Color = Chart_GetColor(i)
    Next i

End Sub


Public Sub SelectedToValue()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectedToValue
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces a targetCell to take on its value.  Removes formulas.
    '---------------------------------------------------------------------------------------
    '
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select the formulas you'd like to convert to static values")

    Dim targetCell As Range
    Dim targetCellValue As String
    For Each targetCell In targetRange
        targetCellValue = targetCell.Value
        targetCell.Clear
        targetCell = targetCellValue
    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No selection made!"
End Sub


Public Sub Selection_ColorWithHex()
    '---------------------------------------------------------------------------------------
    ' Procedure : Selection_ColorWithHex
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Colors a targetCell based on the hex value stored in the targetCell
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetCell As Range
    Dim targetRange As Range
    On Error GoTo errHandler
    Set targetRange = GetInputOrSelection("Select the range of cells to color")
    For Each targetCell In targetRange
        targetCell.Interior.Color = RGB( _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 2, 2)), _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 4, 2)), _
                                    WorksheetFunction.Hex2Dec(Mid(targetCell.Value, 6, 2)))
    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No selection made!"
End Sub


Public Sub SplitAndKeep()
    '---------------------------------------------------------------------------------------
    ' Procedure : SplitAndKeep
    ' Author    : @byronwall
    ' Date      : 2015 08 12
    ' Purpose   : Reduces a targetCell's value to one item returned from Split
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo SplitAndKeep_Error

    Dim rangeToSplit As Range
    Set rangeToSplit = GetInputOrSelection("Select range to split")
    
    If rangeToSplit Is Nothing Then
        Exit Sub
    End If

    Dim delimiter As Variant
    delimiter = InputBox("What delimeter to split on?")
    'StrPtr is undocumented, perhaps add documentation or change function
    If StrPtr(delimiter) = 0 Then
        Exit Sub
    End If

    Dim itemToKeep As Variant
    'Perhaps inform user to input the sequence number of the item to keep
    itemToKeep = InputBox("Which item to keep? (This is 0-indexed)")
    
    If StrPtr(itemToKeep) = 0 Then
        Exit Sub
    End If

    Dim targetCell As Range
    For Each targetCell In Intersect(rangeToSplit, rangeToSplit.Parent.UsedRange)
        
        Dim delimitedCellParts As Variant
        delimitedCellParts = Split(targetCell, delimiter)
        
        If UBound(delimitedCellParts) >= itemToKeep Then
            targetCell.Value = delimitedCellParts(itemToKeep)
        End If
        
    Next targetCell

    On Error GoTo 0
    Exit Sub

SplitAndKeep_Error:
    MsgBox "Check that a valid Range is selected and that a number was entered for which item to keep."
End Sub


Public Sub SplitIntoColumns()
    '---------------------------------------------------------------------------------------
    ' Procedure : SplitIntoColumns
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 24
    ' Purpose   : Splits a targetCell into columns next to it based on a delimeter
    '---------------------------------------------------------------------------------------
    '
    Dim inputRange As Range

    Set inputRange = GetInputOrSelection("Select the range of cells to split")

    Dim targetCell As Range

    Dim delimiter As String
    delimiter = Application.InputBox("What is the delimeter?", , ",", vbOKCancel)
    If delimiter = "" Or delimiter = "False" Then GoTo errHandler
    For Each targetCell In inputRange

        Dim targetCellParts As Variant
        targetCellParts = Split(targetCell, delimiter)

        Dim targetPart As Variant
        For Each targetPart In targetCellParts

            Set targetCell = targetCell.Offset(, 1)
            targetCell = targetPart

        Next targetPart

    Next targetCell
    Exit Sub
errHandler:
    MsgBox "No Delimiter Defined!"
End Sub


Public Sub SplitIntoRows()
    '---------------------------------------------------------------------------------------
    ' Procedure : SplitIntoRows
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Splits a targetCell with return characters into multiple rows with no returns
    '---------------------------------------------------------------------------------------
    '
    Dim outputRange As Range

    Dim inputRange As Range
    Set inputRange = Selection

    Set outputRange = GetInputOrSelection("Select the output corner")

    Dim targetPart As Variant
    Dim offsetCounter As Long
    offsetCounter = 0
    Dim targetCell As Range

    For Each targetCell In inputRange.SpecialCells(xlCellTypeVisible)
        Dim targetParts As Variant
        targetParts = Split(targetCell, vbLf)

        For Each targetPart In targetParts
            outputRange.Offset(offsetCounter) = targetPart

            offsetCounter = offsetCounter + 1
        Next targetPart
    Next targetCell
End Sub


Public Sub TrimSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : TrimSelection
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Trims whitespace from a targetCell's value
    '---------------------------------------------------------------------------------------
    '
    Dim rangeToTrim As Range
    On Error GoTo errHandler
    Set rangeToTrim = GetInputOrSelection("Select the formulas you'd like to convert to static values")

    'disable calcs to speed up
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    'force to only consider used range
    Set rangeToTrim = Intersect(rangeToTrim, rangeToTrim.Parent.UsedRange)

    Dim targetCell As Range
    For Each targetCell In rangeToTrim
        
        'only change if needed
        Dim temporaryTrimHolder As Variant
        temporaryTrimHolder = Trim(targetCell.Value)
        
        'added support for char 160
        'TODO add more characters to remove
        temporaryTrimHolder = Replace(temporaryTrimHolder, Chr(160), vbNullString)
        
        If temporaryTrimHolder <> targetCell.Value Then targetCell.Value = temporaryTrimHolder

    Next targetCell

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub
errHandler:
    MsgBox "No Delimiter Defined!"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

