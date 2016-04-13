Attribute VB_Name = "Usability"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : Usability
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains an assortment of code that automates some task
'---------------------------------------------------------------------------------------


Sub CreatePdfOfEachXlsxFileInFolder()
    
    'pick a folder
    Dim pickedFolder As FileDialog
    Set pickedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    pickedFolder.Show
    
    Dim path As String
    path = pickedFolder.SelectedItems(1) & "\"
    
    'find all files in the folder
    Dim fileName As String
    fileName = Dir(path & "*.xlsx")

    Do While fileName <> ""

        Dim myBook As Workbook
        Set myBook = Workbooks.Open(path & fileName, , True)
        
        Dim mySheet As Worksheet
        
        For Each mySheet In myBook.Worksheets
            mySheet.Range("A16").EntireRow.RowHeight = 15.75
            mySheet.Range("A17").EntireRow.RowHeight = 15.75
            mySheet.Range("A22").EntireRow.RowHeight = 15.75
            mySheet.Range("A23").EntireRow.RowHeight = 15.75
        Next

        myBook.ExportAsFixedFormat xlTypePDF, path & fileName & ".pdf"
        myBook.Close False

        fileName = Dir
    Loop
End Sub

Sub MakeSeveralBoxesWithNumbers()

    Dim myShape As Shape
    Dim mySheet As Worksheet

    Dim selectedRange As Range
    Set selectedRange = Application.InputBox("select range", Type:=8)

    Set mySheet = ActiveSheet

    Dim counter As Long

    For counter = 1 To InputBox("How many?")

        Set myShape = mySheet.Shapes.AddTextbox(msoShapeRectangle, selectedRange.left, _
            selectedRange.top + 20 * counter, 20, 20)

        myShape.title = counter

        myShape.Fill.Visible = msoFalse
        myShape.Line.Visible = msoFalse

        myShape.TextFrame2.TextRange.Characters.Text = counter

        With myShape.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
            .Solid
        End With

    Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ColorInputs
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Finds cells with no value and colors them based on having a formula?
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub ColorInputs()

    Dim myRange As Range
    'This is finding cells that aren't blank, but the description says it should be cells with no values..
    For Each myRange In Selection
        If myRange.Value <> "" Then
            If myRange.HasFormula Then
                myRange.Interior.ThemeColor = msoThemeColorAccent1
            Else
                myRange.Interior.ThemeColor = msoThemeColorAccent2
            End If
        End If
    Next myRange

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CombineAllSheetsData
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Combines all sheets, resuing columns where the same
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub CombineAllSheetsData()

'create the new wkbk and sheet
    Dim combinedBook As Workbook
    Dim dataBook As Workbook

    Set dataBook = ActiveWorkbook
    Set combinedBook = Workbooks.Add

    Dim wsCombined As Worksheet
    Set wsCombined = combinedBook.Sheets.Add

    Dim first As Boolean
    first = True

    Dim combinedRow As Long
    combinedRow = 1

    Dim dataSheet As Worksheet
    For Each dataSheet In dataBook.Sheets
        If dataSheet.name <> wsCombined.name Then

            dataSheet.Unprotect

            'get the headers squared up
            If first Then
                'copy over all headers
                dataSheet.Rows(1).Copy wsCombined.Range("A1")

                first = False
            Else
                'search for missing columns
                Dim headerRange As Range
                For Each headerRange In Intersect(dataSheet.Rows(1), dataSheet.UsedRange)

                    'check if it exists
                    Dim varHdrMatch As Variant
                    varHdrMatch = Application.Match(headerRange, wsCombined.Rows(1), 0)

                    'if not, add to header row
                    If IsError(varHdrMatch) Then
                        wsCombined.Range("A1").End(xlToRight).Offset(, 1) = headerRange
                    End If
                Next headerRange
            End If

            'find the PnPID column for combo
            Dim columnID As Long
            columnID = Application.Match("PnPID", wsCombined.Rows(1), 0)

            'find the PnPID column for data
            Dim dateColumn As Long
            dateColumn = Application.Match("PnPID", dataSheet.Rows(1), 0)

            'add the data, row by row
            Dim myRange As Range
            For Each myRange In dataSheet.UsedRange.SpecialCells(xlCellTypeConstants)
                If myRange.row > 1 Then

                    'check if the PnPID exists in the combo sheet
                    Dim dataRow As Variant
                    dataRow = Application.Match( _
                               dataSheet.Cells(myRange.row, dateColumn), _
                               wsCombined.Columns(columnID), _
                               0)

                    'add new row if it did not exist and id number
                    If IsError(dataRow) Then
                        dataRow = wsCombined.Columns(columnID).Cells(wsCombined.Rows.count, 1).End(xlUp).Offset(1).row
                        wsCombined.Cells(dataRow, columnID) = dataSheet.Cells(myRange.row, dateColumn)
                    End If

                    'get column
                    Dim myColumn As Long
                    myColumn = Application.Match(dataSheet.Cells(1, myRange.Column), wsCombined.Rows(1), 0)

                    'update combo data
                    wsCombined.Cells(dataRow, myColumn) = myRange

                End If
            Next myRange
        End If
    Next dataSheet
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConvertSelectionToCsv
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Crude CSV output from the current selection, works with numbers
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub ConvertSelectionToCsv()

    Dim rangeForCSV As Range
    Set rangeForCSV = GetInputOrSelection("Choose range for converting to CSV")

    If rangeForCSV Is Nothing Then
        Exit Sub
    End If

    Dim outboundCSV As String
    outboundCSV = ""

    Dim csvRow As Range
    For Each csvRow In rangeForCSV.Rows
        
        Dim myArray As Variant
        myArray = Application.Transpose(Application.Transpose(csvRow.Rows.Value2))
        
        'TODO:  improve this to use another Join instead of string concats
        outboundCSV = outboundCSV & Join(myArray, ",") & vbCrLf

    Next csvRow

    Dim clipBoard As MSForms.DataObject
    Set clipBoard = New MSForms.DataObject

    clipBoard.SetText outboundCSV
    clipBoard.PutInClipboard

End Sub

Public Sub CopyCellAddress()
'---------------------------------------------------------------------------------------
' Procedure : CopyCellAddress
' Author    : @byronwall
' Date      : 2015 12 03
' Purpose   : Copies the current cell address to the clipBoard for paste use in a formula
'---------------------------------------------------------------------------------------
'

'TODO: this need to get a button or a keyboard shortcut for easy use
    Dim clipBoard As MSForms.DataObject
    Set clipBoard = New MSForms.DataObject

    Dim selectedRange As Range
    Set selectedRange = Selection

    clipBoard.SetText selectedRange.Address(True, True, xlA1, True)
    clipBoard.PutInClipboard
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CopyClear
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Copies the cells and clears the source Range
'---------------------------------------------------------------------------------------
'
Sub Sheet_DeleteHiddenRows()
    'These rows are unrecoverable
    Dim x As VbMsgBoxResult
    x = MsgBox("This will permanently delete hidden rows. They cannot be recovered. Are you sure?", vbYesNo)
    
    If Not x = vbYes Then
        Exit Sub
    End If
        
    Application.ScreenUpdating = False
    
    'We might as well tell the user how many rows were hidden
    Dim counter As Long
    counter = 0
    With ActiveSheet
        Dim i As Long
        For i = .UsedRange.Rows.count To 1 Step -1
            If .Rows(i).Hidden Then
                .Rows(i).Delete
                counter = counter + 1
            End If
        Next i
    End With
    Application.ScreenUpdating = True
    
    MsgBox (counter & " rows were deleted")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CutPasteTranspose
' Author    : @byronwall, @RaymondWise
' Date      : 2015 07 31
' Purpose   : Does a cut/transpose by cutting each cell individually
'---------------------------------------------------------------------------------------
'

'########Still Needs to address Issue#23#############
Sub CutPasteTranspose()

    On Error GoTo errHandler
    Dim rangeToSelect As Range
    'TODO #Should use new inputbox function
    Set rangeToSelect = Selection

    Dim outboundRange As Range
    Set outboundRange = Application.InputBox("Select output corner", Type:=8)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual



    Dim myCorner As Range
    Set myCorner = rangeToSelect.Cells(1, 1)

    Dim myRow As Long
    myRow = myCorner.row
    Dim incomingColumn As Long
    incomingColumn = myCorner.Column

    Dim outboundRow As Long
    Dim outboundColumn As Long
    outboundRow = outboundRange.row
    outboundColumn = outboundRange.Column

    outboundRange.Activate
    
    'Check to not overwrite
    Dim myRange As Range
    For Each myRange In rangeToSelect
        If Not Intersect(rangeToSelect, Cells(outboundRow + myRange.Column - incomingColumn, outboundColumn + myRange.row - myRow)) Is Nothing Then
            MsgBox ("Your destination intersects with your data")
            Exit Sub
        End If
    Next
    
    For Each myRange In rangeToSelect
        myRange.Cut
        ActiveSheet.Cells(outboundRow + myRange.Column - incomingColumn, outboundColumn + myRange.row - myRow).Activate
        ActiveSheet.Paste
    Next myRange

    Application.CutCopyMode = False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
errHandler:
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EvaluateArrayFormulaOnNewSheet
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Wacky thing to force an Array formula to return as an Array
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub EvaluateArrayFormulaOnNewSheet()

'cut cell with formula
    Dim streetAddress As String
    Dim startingRange As Range
    Set startingRange = Sheet1.Range("J2")
    streetAddress = startingRange.Address

    startingRange.Cut

    'create new sheet
    Dim mySheet As Worksheet
    Set mySheet = Worksheets.Add

    'paste cell onto sheet
    Dim myArrayRange As Range
    Set myArrayRange = mySheet.Range("A1")
    mySheet.Paste myArrayRange

    'expand Array formula size.. resize to whatever size is needed
    myArrayRange.Resize(3).FormulaArray = myArrayRange.FormulaArray

    'get your result
    Dim myArray As Variant
    myArray = Application.Evaluate(myArrayRange.CurrentArray.Address)

    ''''do something with your result here... it is an Array


    'shrink the formula back to one cell
    Dim myFormula As String
    myFormula = myArrayRange.FormulaArray

    myArrayRange.CurrentArray.ClearContents
    myArrayRange.FormulaArray = myFormula

    'cut and paste back to original spot
    myArrayRange.Cut

    Sheet1.Paste Sheet1.Range(streetAddress)

    Application.DisplayAlerts = False
    mySheet.Delete
    Application.DisplayAlerts = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportFilesFromFolder
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Goes through a folder and process all workbooks therein
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub ExportFilesFromFolder()
    '###Needs error handling
    'TODO: consider deleting this Sub since it is quite specific
    Application.ScreenUpdating = False

    Dim file As Variant
    Dim path As String
    path = InputBox("What path?")

    file = Dir(path)
    While (file <> "")

        Debug.Print path & file

        Dim fileName As String

        fileName = path & file

        Dim wbActive As Workbook
        Set wbActive = Workbooks.Open(fileName)

        Dim wsActive As Worksheet
        Set wsActive = wbActive.Sheets("Case Summary")

        With ActiveSheet.PageSetup
            .TopMargin = Application.InchesToPoints(0.4)
            .BottomMargin = Application.InchesToPoints(0.4)
        End With

        wsActive.ExportAsFixedFormat xlTypePDF, path & "PDFs\" & file & ".pdf"

        wbActive.Close False

        file = Dir
    Wend

    Application.ScreenUpdating = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FillValueDown
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Does a fill of blank values from the cell above with a value
'---------------------------------------------------------------------------------------
'
Sub FillValueDown()

    Dim inputRange As Range
    Set inputRange = GetInputOrSelection("Select range for waterfall")

    If inputRange Is Nothing Then
        Exit Sub
    End If

    Dim myRange As Range
    For Each myRange In Intersect(inputRange.SpecialCells(xlCellTypeBlanks), inputRange.Parent.UsedRange)
        myRange = myRange.End(xlUp)
    Next myRange

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ForceRecalc
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Provides a button to do a full recalc
'---------------------------------------------------------------------------------------
'
Sub ForceRecalc()

    Application.CalculateFullRebuild

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GenerateRandomData
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Generates a block of random data for testing questions on SO
'---------------------------------------------------------------------------------------
'
Sub GenerateRandomData()

    Dim myRange As Range
    Set myRange = Range("B2")

    Dim i As Long

    For i = 0 To 3
        myRange.Offset(, i) = Chr(65 + i)

        With myRange.Offset(1, i).Resize(10)
            Select Case i
            Case 0
                .Formula = "=TODAY()+ROW()"
            Case Else
                .Formula = "=RANDBETWEEN(1,100)"
            End Select

            .Value = .Value
        End With
    Next i

    ActiveSheet.UsedRange.Columns.ColumnWidth = 15

End Sub

'---------------------------------------------------------------------------------------
' Procedure : OpenContainingFolder
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Open the folder that contains the ActiveWorkbook
'---------------------------------------------------------------------------------------
'
Sub OpenContainingFolder()

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb.path <> "" Then
        wb.FollowHyperlink wb.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PivotSetAllFields
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets all fields in a PivotTable to use a certain calculation type
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub PivotSetAllFields()

    Dim myPivotTable As PivotTable
    Dim mySheet As Worksheet

    Set mySheet = ActiveSheet

    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."

    For Each myPivotTable In mySheet.PivotTables
        Dim myPivotField As PivotField
        For Each myPivotField In myPivotTable.DataFields
            On Error Resume Next
            myPivotField.Function = xlAverage
        Next myPivotField
    Next myPivotTable

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SeriesSplit
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Takes a category columns and splits the values out into new columns for each unique entry
'---------------------------------------------------------------------------------------
'
Sub SeriesSplit()

    On Error GoTo ErrorNoSelection

    Dim rangeToSelection As Range
    Set rangeToSelection = Application.InputBox("Select category range with heading", Type:=8)
    Set rangeToSelection = Intersect(rangeToSelection, rangeToSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    On Error GoTo 0

    'determine default value
    Dim defaultValue As Variant
    defaultValue = InputBox("Enter the default value", , "#N/A")

    'detect cancel and exit
    If StrPtr(defaultValue) = 0 Then
        Exit Sub
    End If

    Dim dictCategories As New Dictionary

    Dim myCategory As Range
    For Each myCategory In rangeToSelection
        'skip the header row
        If myCategory.Address <> rangeToSelection.Cells(1).Address Then
            dictCategories(myCategory.Value) = 1
        End If

    Next myCategory

    valueRange.EntireColumn.Offset(, 1).Resize(, dictCategories.count).Insert
    'head the columns with the values

    Dim myValues As Variant
    Dim counter As Long
    counter = 1
    For Each myValues In dictCategories
        valueRange.Cells(1).Offset(, counter) = myValues
        counter = counter + 1
    Next myValues

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim myFormula As Variant
    myFormula = "=IF(RC" & rangeToSelection.Column & " =R" & _
                 valueRange.Cells(1).row & "C,RC" & valueRange.Column & "," & defaultValue & ")"

    Dim rngFormula As Range
    Set rngFormula = valueRange.Offset(1, 1).Resize(valueRange.Rows.count - 1, dictCategories.count)
    rngFormula.FormulaR1C1 = myFormula
    rngFormula.EntireColumn.AutoFit

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SeriesSplitIntoBins
' Author    : @byronwall
' Date      : 2015 11 03
' Purpose   : Code will break a column of continuous data into bins for plotting
'---------------------------------------------------------------------------------------
'
Sub SeriesSplitIntoBins()

    On Error GoTo ErrorNoSelection

    Dim rangeToSelection As Range
    Set rangeToSelection = Application.InputBox("Select category range with heading", _
                                            Type:=8)
    Set rangeToSelection = Intersect(rangeToSelection, _
                                 rangeToSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + _
                                                                                                xlNumbers + xlTextValues)

    Dim valueRange As Range
    Set valueRange = Application.InputBox("Select values range with heading", _
                                         Type:=8)
    Set valueRange = Intersect(valueRange, valueRange.Parent.UsedRange)

    ''need to prompt for max/min/bins
    Dim maximumValue As Double, minimumValue As Double, myBins As Long

    minimumValue = Application.InputBox("Minimum value.", "Min", _
                                   WorksheetFunction.Min(rangeToSelection), Type:=1)
    maximumValue = Application.InputBox("Maximum value.", "Max", _
                                   WorksheetFunction.Max(rangeToSelection), Type:=1)
    myBins = Application.InputBox("Number of groups.", "Bins", _
                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.count(rangeToSelection)), _
                                                                0), Type:=1)

    On Error GoTo 0

    'determine default value
    Dim defaultValue As Variant
    defaultValue = Application.InputBox("Enter the default value", "Default", _
                                      "#N/A")

    'detect cancel and exit
    If StrPtr(defaultValue) = 0 Then
        Exit Sub
    End If

    ''TODO prompt for output location

    valueRange.EntireColumn.Offset(, 1).Resize(, myBins + 2).Insert
    'head the columns with the values

    ''TODO add a For loop to go through the bins

    Dim binNumber As Long
    For binNumber = 0 To myBins
        valueRange.Cells(1).Offset(, binNumber + 1) = minimumValue + (maximumValue - _
                                                                minimumValue) * binNumber / myBins
    Next

    'add the last item
    valueRange.Cells(1).Offset(, myBins + 2).FormulaR1C1 = "=RC[-1]"

    'FIRST =IF($D2 <=V$1,$U2,#N/A)
    '=IF(RC4 <=R1C,RC21,#N/A)

    'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  '''W current, then left
    '=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)

    'LAST =IF($D2>AA$1,$U2,#N/A)
    '=IF(RC4>R1C[-1],RC21,#N/A)

    ''TODO add number format to display header correctly (helps with charts)

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim myFormula As Variant
    myFormula = "=IF(AND(RC" & rangeToSelection.Column & " <=R" & _
                 valueRange.Cells(1).row & "C," & "RC" & rangeToSelection.Column & ">R" & _
                 valueRange.Cells(1).row & "C[-1]" & ")" & ",RC" & valueRange.Column & "," & _
                 defaultValue & ")"

    Dim firstFormula As Variant
    firstFormula = "=IF(AND(RC" & rangeToSelection.Column & " <=R" & _
                       valueRange.Cells(1).row & "C)" & ",RC" & valueRange.Column & "," & defaultValue _
                     & ")"

    Dim lastFormula As Variant
    lastFormula = "=IF(AND(RC" & rangeToSelection.Column & " >R" & _
                      valueRange.Cells(1).row & "C)" & ",RC" & valueRange.Column & "," & defaultValue _
                    & ")"

    Dim rngFormula As Range
    Set rngFormula = valueRange.Offset(1, 1).Resize(valueRange.Rows.count - 1, _
                                                   myBins + 2)
    rngFormula.FormulaR1C1 = myFormula

    'override with first/last
    rngFormula.Columns(1).FormulaR1C1 = firstFormula
    rngFormula.Columns(rngFormula.Columns.count).FormulaR1C1 = lastFormula

    rngFormula.EntireColumn.AutoFit

    'set the number formats
    rngFormula.Offset(-1).Rows(1).Resize(1, myBins + 1).NumberFormat = _
    "<= General"
    rngFormula.Offset(-1).Rows(1).Offset(, myBins + 1).NumberFormat = _
    "> General"

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : mySheet_DeleteHiddenRows
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Deletes the hidden rows in a sheet.  Good for a "permanent" filter
'---------------------------------------------------------------------------------------
'
'Changed sub name to avoid reserved object name
Sub mySheet_DeleteHiddenRows()

    Application.ScreenUpdating = False
    Dim row As Range
    Dim i As Long
    For i = ActiveSheet.UsedRange.Rows.count To 1 Step -1


        Set row = ActiveSheet.Rows(i)

        If row.Hidden Then
            row.Delete
        End If
    Next i

    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : UnhideAllRowsAndColumns
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Unhides everything in a Worksheet
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub UnhideAllRowsAndColumns()

    ActiveSheet.Cells.EntireRow.Hidden = False
    ActiveSheet.Cells.EntireColumn.Hidden = False

End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateScrollbars
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Cheap trick that forces Excel to update the scroll bars after a large deletion
'---------------------------------------------------------------------------------------
'
Sub UpdateScrollbars()

    Dim myRange As Variant
    myRange = ActiveSheet.UsedRange.Address

End Sub

