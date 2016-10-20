Attribute VB_Name = "Usability"
Option Explicit

Public Sub ColorInputs()
    '---------------------------------------------------------------------------------------
    ' Procedure : ColorInputs
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Finds cells with no value and colors them based on having a formula?
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetCell As Range
    Const FIRST_COLOR_ACCENT As String = "msoThemeColorAccent1"
    Const SECOND_COLOR_ACCENT As String = "msoThemeColorAccent2"
    'This is finding cells that aren't blank, but the description says it should be cells with no values..
    For Each targetCell In Selection
        If targetCell.Value <> "" Then
            If targetCell.HasFormula Then
                targetCell.Interior.ThemeColor = FIRST_COLOR_ACCENT
            Else
                targetCell.Interior.ThemeColor = SECOND_COLOR_ACCENT
            End If
        End If
    Next targetCell

End Sub


Public Sub CombineAllSheetsData()
    '---------------------------------------------------------------------------------------
    ' Procedure : CombineAllSheetsData
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Combines all sheets, resuing columns where the same
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'create the new wkbk and sheet
    Dim targetWorkbook As Workbook
    Dim sourceWorkbook As Workbook

    Set sourceWorkbook = ActiveWorkbook
    Set targetWorkbook = Workbooks.Add

    Dim targetWorksheet As Worksheet
    Set targetWorksheet = targetWorkbook.Sheets.Add

    Dim isFirst As Boolean
    isFirst = True

    Dim targetRow As Long
    targetRow = 1

    Dim sourceWorksheet As Worksheet
    For Each sourceWorksheet In sourceWorkbook.Sheets
        If sourceWorksheet.name <> targetWorksheet.name Then

            sourceWorksheet.Unprotect

            'get the headers squared up
            If isFirst Then
                'copy over all headers
                sourceWorksheet.Rows(1).Copy targetWorksheet.Range("A1")
                isFirst = False
            
            Else
                'search for missing columns
                Dim headerRow As Range
                For Each headerRow In Intersect(sourceWorksheet.Rows(1), sourceWorksheet.UsedRange)

                    'check if it exists
                    Dim matchingHeader As Variant
                    matchingHeader = Application.Match(headerRow, targetWorksheet.Rows(1), 0)

                    'if not, add to header row
                    If IsError(matchingHeader) Then targetWorksheet.Range("A1").End(xlToRight).Offset(, 1) = headerRow
                Next headerRow
            End If

            'find the PnPID column for combo
            Dim pIDColumn As Long
            pIDColumn = Application.Match("PnPID", targetWorksheet.Rows(1), 0)

            'find the PnPID column for data
            Dim pIDData As Long
            pIDData = Application.Match("PnPID", sourceWorksheet.Rows(1), 0)

            'add the data, row by row
            Dim targetCell As Range
            For Each targetCell In sourceWorksheet.UsedRange.SpecialCells(xlCellTypeConstants)
                If targetCell.Row > 1 Then

                    'check if the PnPID exists in the combo sheet
                    Dim sourceRow As Variant
                    sourceRow = Application.Match( _
                               sourceWorksheet.Cells(targetCell.Row, pIDData), _
                               targetWorksheet.Columns(pIDColumn), _
                               0)

                    'add new row if it did not exist and id number
                    If IsError(sourceRow) Then
                        sourceRow = targetWorksheet.Columns(pIDColumn).Cells(targetWorksheet.Rows.count, 1).End(xlUp).Offset(1).Row
                        targetWorksheet.Cells(sourceRow, pIDColumn) = sourceWorksheet.Cells(targetCell.Row, pIDData)
                    End If

                    'get column
                    Dim columnNumber As Long
                    columnNumber = Application.Match(sourceWorksheet.Cells(1, targetCell.Column), targetWorksheet.Rows(1), 0)

                    'update combo data
                    targetWorksheet.Cells(sourceRow, columnNumber) = targetCell

                End If
            Next targetCell
        End If
    Next sourceWorksheet
End Sub


Public Sub ConvertSelectionToCsv()
    '---------------------------------------------------------------------------------------
    ' Procedure : ConvertSelectionToCsv
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Crude CSV output from the current selection, works with numbers
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim sourceRange As Range
    Set sourceRange = GetInputOrSelection("Choose range for converting to CSV")

    If sourceRange Is Nothing Then Exit Sub

    Dim outputString As String

    Dim dataRow As Range
    For Each dataRow In sourceRange.Rows
        
        Dim dataArray As Variant
        dataArray = Application.Transpose(Application.Transpose(dataRow.Rows.Value2))
        
        'TODO:  improve this to use another Join instead of string concats
        outputString = outputString & Join(dataArray, ",") & vbCrLf

    Next dataRow

    Dim myClipboard As MSForms.DataObject
    Set myClipboard = New MSForms.DataObject

    myClipboard.SetText outputString
    myClipboard.PutInClipboard

End Sub

Public Sub CopyCellAddress()
    '---------------------------------------------------------------------------------------
    ' Procedure : CopyCellAddress
    ' Author    : @byronwall
    ' Date      : 2015 12 03
    ' Purpose   : Copies the current cell address to the myClipboard for paste use in a formula
    '---------------------------------------------------------------------------------------
    '

    'TODO: this need to get a button or a keyboard shortcut for easy use
    Dim myClipboard As MSForms.DataObject
    Set myClipboard = New MSForms.DataObject

    Dim sourceRange As Range
    Set sourceRange = Selection

    myClipboard.SetText sourceRange.Address(True, True, xlA1, True)
    myClipboard.PutInClipboard
End Sub

Public Sub Sheet_DeleteHiddenRows()
    'These rows are unrecoverable
    Dim shouldDeleteHiddenRows As VbMsgBoxResult
    shouldDeleteHiddenRows = MsgBox("This will permanently delete hidden rows. They cannot be recovered. Are you sure?", vbYesNo)
    
    If Not shouldDeleteHiddenRows = vbYes Then Exit Sub
        
    Application.ScreenUpdating = False
    
    'collect a range to delete at end, using UNION-DELETE
    Dim rangeToDelete As Range
    
    Dim counter As Long
    counter = 0
    With ActiveSheet
        Dim rowIndex As Long
        For rowIndex = .UsedRange.Rows.count To 1 Step -1
            If .Rows(rowIndex).Hidden Then
                If rangeToDelete Is Nothing Then
                    Set rangeToDelete = .Rows(rowIndex)
                Else
                    Set rangeToDelete = Union(rangeToDelete, .Rows(rowIndex))
                End If
                counter = counter + 1
            End If
        Next rowIndex
    End With
    
    rangeToDelete.Delete
    
    Application.ScreenUpdating = True
    
    MsgBox (counter & " rows were deleted")
End Sub


Public Sub CutPasteTranspose()
    '---------------------------------------------------------------------------------------
    ' Procedure : CutPasteTranspose
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 31
    ' Purpose   : Does a cut/transpose by cutting each cell individually
    '---------------------------------------------------------------------------------------
    '

    '########Still Needs to address Issue#23#############
    On Error GoTo errHandler
    Dim sourceRange As Range
    'TODO #Should use new inputbox function
    Set sourceRange = Selection

    Dim outputRange As Range
    Set outputRange = Application.InputBox("Select output corner", Type:=8)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim topLeftCell As Range
    Set topLeftCell = sourceRange.Cells(1, 1)

    Dim topRow As Long
    topRow = topLeftCell.Row
    Dim leftColumn As Long
    leftColumn = topLeftCell.Column

    Dim outputRow As Long
    Dim outputColumn As Long
    outputRow = outputRange.Row
    outputColumn = outputRange.Column

    outputRange.Activate
    
    'Check to not overwrite
    Dim targetCell As Range
    For Each targetCell In sourceRange
        If Not Intersect(sourceRange, Cells(outputRow + targetCell.Column - leftColumn, outputColumn + targetCell.Row - topRow)) Is Nothing Then
            MsgBox ("Your destination intersects with your data. Exiting.")
            GoTo errHandler
        End If
    Next
    
    'this can be better
    For Each targetCell In sourceRange
        targetCell.Cut
        ActiveSheet.Cells(outputRow + targetCell.Column - leftColumn, outputColumn + targetCell.Row - topRow).Activate
        ActiveSheet.Paste
    Next targetCell
    
errHandler:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

End Sub

Public Sub FillValueDown()
    '---------------------------------------------------------------------------------------
    ' Procedure : FillValueDown
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Does a fill of blank values from the cell above with a value
    '---------------------------------------------------------------------------------------
    '
    Dim inputRange As Range
    Set inputRange = GetInputOrSelection("Select range for waterfall")

    If inputRange Is Nothing Then Exit Sub

    Dim targetCell As Range
    For Each targetCell In Intersect(inputRange.SpecialCells(xlCellTypeBlanks), inputRange.Parent.UsedRange)
        targetCell = targetCell.End(xlUp)
    Next targetCell

End Sub


Public Sub ForceRecalc()
    '---------------------------------------------------------------------------------------
    ' Procedure : ForceRecalc
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Provides a button to do a full recalc
    '---------------------------------------------------------------------------------------
    '
    Application.CalculateFullRebuild

End Sub


Public Sub GenerateRandomData()
    '---------------------------------------------------------------------------------------
    ' Procedure : GenerateRandomData
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Generates a block of random data for testing questions on SO
    ' Description: Will create a data table of 4 columns and fill the first column with dates and the others with integers
    '---------------------------------------------------------------------------------------
    '
    Const NUMBER_OF_ROWS As Long = 10
    Const NUMBER_OF_COLUMNS As Long = 3 '0 index
    Const DEFAULT_COLUMN_WIDTH As Long = 15
    
    'Since we only work with offset, targetcell can be a constant, but range constants are awkward
    Dim targetCell As Range
    Set targetCell = Range("B2")

    Dim i As Long

    For i = 0 To NUMBER_OF_COLUMNS
        targetCell.Offset(, i) = Chr(65 + i)

        With targetCell.Offset(1, i).Resize(NUMBER_OF_ROWS)
            Select Case i
            Case 0
                .Formula = "=TODAY()+ROW()"
            Case Else
                .Formula = "=RANDBETWEEN(1,100)"
            End Select

            .Value = .Value
        End With
    Next i

    ActiveSheet.UsedRange.Columns.ColumnWidth = DEFAULT_COLUMN_WIDTH

End Sub


Public Sub OpenContainingFolder()
    '---------------------------------------------------------------------------------------
    ' Procedure : OpenContainingFolder
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Open the folder that contains the ActiveWorkbook
    '---------------------------------------------------------------------------------------
    '
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    If targetWorkbook.path <> "" Then
        targetWorkbook.FollowHyperlink targetWorkbook.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub


Public Sub PivotSetAllFields()
    '---------------------------------------------------------------------------------------
    ' Procedure : PivotSetAllFields
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets all fields in a PivotTable to use a certain calculation type
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetTable As PivotTable
    Dim targetSheet As Worksheet

    Set targetSheet = ActiveSheet

    'this information is a bit unclear to me
    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."
    On Error Resume Next
    For Each targetTable In targetSheet.PivotTables
        Dim targetField As PivotField
        For Each targetField In targetTable.DataFields
            targetField.Function = xlAverage
        Next targetField
    Next targetTable

End Sub

Public Sub SeriesSplit()
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesSplit
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Takes a category columns and splits the values out into new columns for each unique entry
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo ErrorNoSelection

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("Select category range with heading", Type:=8)
    Set selectedRange = Intersect(selectedRange, selectedRange.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    On Error GoTo 0

    'determine default value
    Dim defaultString As Variant
    defaultString = InputBox("Enter the default value", , "#N/A")
    'strptr is undocumented
    'detect cancel and exit
    If StrPtr(defaultString) = 0 Then
        Exit Sub
    End If

    Dim dictCategories As New Dictionary

    Dim categoryRange As Range
    For Each categoryRange In selectedRange
        'skip the header row
        If categoryRange.Address <> selectedRange.Cells(1).Address Then dictCategories(categoryRange.Value) = 1
    Next categoryRange

    valueRange.EntireColumn.Offset(, 1).Resize(, dictCategories.count).Insert
    'head the columns with the values

    Dim valueCollection As Variant
    Dim counter As Long
    counter = 1
    For Each valueCollection In dictCategories
        valueRange.Cells(1).Offset(, counter) = valueCollection
        counter = counter + 1
    Next valueCollection

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim formulaHolder As Variant
    formulaHolder = "=IF(RC" & selectedRange.Column & " =R" & _
                 valueRange.Cells(1).Row & "C,RC" & valueRange.Column & "," & defaultString & ")"

    Dim formulaRange As Range
    Set formulaRange = valueRange.Offset(1, 1).Resize(valueRange.Rows.count - 1, dictCategories.count)
    formulaRange.FormulaR1C1 = formulaHolder
    formulaRange.EntireColumn.AutoFit

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub


Public Sub SeriesSplitIntoBins()
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesSplitIntoBins
    ' Author    : @byronwall
    ' Date      : 2015 11 03
    ' Purpose   : Code will break a column of continuous data into bins for plotting
    '---------------------------------------------------------------------------------------
    '
    Const LESS_THAN_EQUAL_TO_GENERAL As String = "<= General"
    Const GREATER_THAN_GENERAL As String = "> General"
    On Error GoTo ErrorNoSelection

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("Select category range with heading", Type:=8)
    Set selectedRange = Intersect(selectedRange, selectedRange.Parent.UsedRange) _
                                 .SpecialCells(xlCellTypeVisible, xlLogical + _
                                  xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    ''need to prompt for max/min/bins
    Dim maximumValue As Double, minimumValue As Double, binValue As Long

    minimumValue = Application.InputBox("Minimum value.", "Min", _
                                        WorksheetFunction.Min(selectedRange), Type:=1)
                                   
    maximumValue = Application.InputBox("Maximum value.", "Max", _
                                        WorksheetFunction.Max(selectedRange), Type:=1)
                                   
    binValue = Application.InputBox("Number of groups.", "Bins", _
                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.count(selectedRange)), _
                                    0), Type:=1)

    On Error GoTo 0

    'determine default value
    Dim defaultString As Variant
    defaultString = Application.InputBox("Enter the default value", "Default", "#N/A")

    'detect cancel and exit
    If StrPtr(defaultString) = 0 Then Exit Sub

    ''TODO prompt for output location

    valueRange.EntireColumn.Offset(, 1).Resize(, binValue + 2).Insert
    'head the columns with the values

    ''TODO add a For loop to go through the bins

    Dim targetBin As Long
    For targetBin = 0 To binValue
        valueRange.Cells(1).Offset(, targetBin + 1) = minimumValue + (maximumValue - _
                                                      minimumValue) * targetBin / binValue
    Next

    'add the last item
    valueRange.Cells(1).Offset(, binValue + 2).FormulaR1C1 = "=RC[-1]"

    'FIRST =IF($D2 <=V$1,$U2,#N/A)
    '=IF(RC4 <=R1C,RC21,#N/A)

    'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  '''W current, then left
    '=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)

    'LAST =IF($D2>AA$1,$U2,#N/A)
    '=IF(RC4>R1C[-1],RC21,#N/A)

    ''TODO add number format to display header correctly (helps with charts)

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim formulaHolder As Variant
    formulaHolder = "=IF(AND(RC" & selectedRange.Column & " <=R" & _
                    valueRange.Cells(1).Row & "C," & "RC" & selectedRange.Column & ">R" & _
                    valueRange.Cells(1).Row & "C[-1]" & ")" & ",RC" & valueRange.Column & "," & _
                    defaultString & ")"

    Dim firstFormula As Variant
    firstFormula = "=IF(AND(RC" & selectedRange.Column & " <=R" & _
                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _
                    & ")"

    Dim lastFormula As Variant
    lastFormula = "=IF(AND(RC" & selectedRange.Column & " >R" & _
                    valueRange.Cells(1).Row & "C)" & ",RC" & valueRange.Column & "," & defaultString _
                    & ")"

    Dim formulaRange As Range
    Set formulaRange = valueRange.Offset(1, 1).Resize(valueRange.Rows.count - 1, binValue + 2)
    formulaRange.FormulaR1C1 = formulaHolder

    'override with first/last
    formulaRange.Columns(1).FormulaR1C1 = firstFormula
    formulaRange.Columns(formulaRange.Columns.count).FormulaR1C1 = lastFormula

    formulaRange.EntireColumn.AutoFit

    'set the number formats

    formulaRange.Offset(-1).Rows(1).Resize(1, binValue + 1).NumberFormat = LESS_THAN_EQUAL_TO_GENERAL
    formulaRange.Offset(-1).Rows(1).Offset(, binValue + 1).NumberFormat = GREATER_THAN_GENERAL

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub


Public Sub UnhideAllRowsAndColumns()
    '---------------------------------------------------------------------------------------
    ' Procedure : UnhideAllRowsAndColumns
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Unhides everything in a Worksheet
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    ActiveSheet.Cells.EntireRow.Hidden = False
    ActiveSheet.Cells.EntireColumn.Hidden = False

End Sub


Public Sub UpdateScrollbars()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateScrollbars
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Cheap trick that forces Excel to update the scroll bars after a large deletion
    '---------------------------------------------------------------------------------------
    '
    Dim targetRange As Variant
    targetRange = ActiveSheet.UsedRange.Address

End Sub

