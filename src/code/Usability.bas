Attribute VB_Name = "Usability"
'---------------------------------------------------------------------------------------
' Module    : Usability
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains an assortment of code that automates some task
'---------------------------------------------------------------------------------------


Sub CreatePdfOfEachXlsxFileInFolder()
    
    'pick a folder
    Dim diag_folder As FileDialog
    Set diag_folder = Application.FileDialog(msoFileDialogFolderPicker)
    
    diag_folder.Show
    
    Dim str_path As String
    str_path = diag_folder.SelectedItems(1) & "\"
    
    'find all files in the folder
    Dim str_file As String
    str_file = Dir(str_path & "*.xlsx")

    Do While str_file <> ""

        Dim wkbk_file As Workbook
        Set wkbk_file = Workbooks.Open(str_path & str_file, , True)
        
        Dim sht As Worksheet
        
        For Each sht In wkbk_file.Worksheets
            sht.Range("A16").EntireRow.RowHeight = 15.75
            sht.Range("A17").EntireRow.RowHeight = 15.75
            sht.Range("A22").EntireRow.RowHeight = 15.75
            sht.Range("A23").EntireRow.RowHeight = 15.75
        Next

        wkbk_file.ExportAsFixedFormat xlTypePDF, str_path & str_file & ".pdf"
        wkbk_file.Close False

        str_file = Dir
    Loop
End Sub

Sub MakeSeveralBoxesWithNumbers()

    Dim shp As Shape
    Dim sht As Worksheet

    Dim rng_loc As Range
    Set rng_loc = Application.InputBox("select range", Type:=8)

    Set sht = ActiveSheet

    Dim int_counter As Integer

    For int_counter = 1 To InputBox("How many?")

        Set shp = sht.Shapes.AddTextbox(msoShapeRectangle, rng_loc.left, _
            rng_loc.top + 20 * int_counter, 20, 20)

        shp.Title = int_counter

        shp.Fill.Visible = msoFalse
        shp.Line.Visible = msoFalse

        shp.TextFrame2.TextRange.Characters.Text = int_counter

        With shp.TextFrame2.TextRange.Font.Fill
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

    Dim c As Range
    'This is finding cells that aren't blank, but the description says it should be cells with no values..
    For Each c In Selection
        If c.Value <> "" Then
            If c.HasFormula Then
                c.Interior.ThemeColor = msoThemeColorAccent1
            Else
                c.Interior.ThemeColor = msoThemeColorAccent2
            End If
        End If
    Next c

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
    Dim wbCombo As Workbook
    Dim wbData As Workbook

    Set wbData = ActiveWorkbook
    Set wbCombo = Workbooks.Add

    Dim wsCombined As Worksheet
    Set wsCombined = wbCombo.Sheets.Add

    Dim boolFirst As Boolean
    boolFirst = True

    Dim iComboRow As Integer
    iComboRow = 1

    Dim wsData As Worksheet
    For Each wsData In wbData.Sheets
        If wsData.name <> wsCombined.name Then

            wsData.Unprotect

            'get the headers squared up
            If boolFirst Then
                'copy over all headers
                wsData.Rows(1).Copy wsCombined.Range("A1")

                boolFirst = False
            Else
                'search for missing columns
                Dim rngHeader As Range
                For Each rngHeader In Intersect(wsData.Rows(1), wsData.UsedRange)

                    'check if it exists
                    Dim varHdrMatch As Variant
                    varHdrMatch = Application.Match(rngHeader, wsCombined.Rows(1), 0)

                    'if not, add to header row
                    If IsError(varHdrMatch) Then
                        wsCombined.Range("A1").End(xlToRight).Offset(, 1) = rngHeader
                    End If
                Next rngHeader
            End If

            'find the PnPID column for combo
            Dim int_colId As Integer
            int_colId = Application.Match("PnPID", wsCombined.Rows(1), 0)

            'find the PnPID column for data
            Dim iColIDData As Integer
            iColIDData = Application.Match("PnPID", wsData.Rows(1), 0)

            'add the data, row by row
            Dim c As Range
            For Each c In wsData.UsedRange.SpecialCells(xlCellTypeConstants)
                If c.Row > 1 Then

                    'check if the PnPID exists in the combo sheet
                    Dim iDataRow As Variant
                    iDataRow = Application.Match( _
                               wsData.Cells(c.Row, iColIDData), _
                               wsCombined.Columns(int_colId), _
                               0)

                    'add new row if it did not exist and id number
                    If IsError(iDataRow) Then
                        iDataRow = wsCombined.Columns(int_colId).Cells(wsCombined.Rows.count, 1).End(xlUp).Offset(1).Row
                        wsCombined.Cells(iDataRow, int_colId) = wsData.Cells(c.Row, iColIDData)
                    End If

                    'get column
                    Dim iCol As Integer
                    iCol = Application.Match(wsData.Cells(1, c.Column), wsCombined.Rows(1), 0)

                    'update combo data
                    wsCombined.Cells(iDataRow, iCol) = c

                End If
            Next c
        End If
    Next wsData
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

    Dim rngCSV As Range
    Set rngCSV = GetInputOrSelection

    If rngCSV Is Nothing Then
        Exit Sub
    End If

    Dim csvOut As String
    csvOut = ""

    Dim csvRow As Range
    For Each csvRow In rngCSV.Rows
        
        Dim arr As Variant
        arr = Application.Transpose(Application.Transpose(csvRow.Rows.Value2))
        
        'TODO:  improve this to use another Join instead of string concats
        csvOut = csvOut & Join(arr, ",") & vbCrLf

    Next csvRow

    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    clipboard.SetText csvOut
    clipboard.PutInClipboard

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
    x = MsgBox("This will permanently delete hidden rows. They cannot be recovered. Are you sure?", vbYesNo)
        If x = 7 Then Exit Sub
        
    Application.ScreenUpdating = False
    
    'We might as well tell the user how many rows were hidden
    Dim iCount As Integer
    iCount = 0
    With ActiveSheet
        For i = .UsedRange.Rows.count To 1 Step -1
            If .Rows(i).Hidden Then
                .Rows(i).Delete
                iCount = iCount + 1
            End If
        Next i
    End With
    Application.ScreenUpdating = True
    
    MsgBox (iCount & " rows were deleted")
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
    Dim rngSelect As Range
    'TODO #Should use new inputbox function
    Set rngSelect = Selection

    Dim rngOut As Range
    Set rngOut = Application.InputBox("Select output corner", Type:=8)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual



    Dim rCorner As Range
    Set rCorner = rngSelect.Cells(1, 1)

    Dim iCRow As Integer
    iCRow = rCorner.Row
    Dim iCCol As Integer
    iCCol = rCorner.Column

    Dim iORow As Integer
    Dim iOCol As Integer
    iORow = rngOut.Row
    iOCol = rngOut.Column

    rngOut.Activate
    
    'Check to not overwrite
    For Each c In rngSelect
        If Not Intersect(rngSelect, Cells(iORow + c.Column - iCCol, iOCol + c.Row - iCRow)) Is Nothing Then
            MsgBox ("Your destination intersects with your data")
            Exit Sub
        End If
    Next
    
    For Each c In rngSelect
        c.Cut
        ActiveSheet.Cells(iORow + c.Column - iCCol, iOCol + c.Row - iCRow).Activate
        ActiveSheet.Paste
    Next c

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
' Purpose   : Wacky thing to force an array formula to return as an array
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub EvaluateArrayFormulaOnNewSheet()

'cut cell with formula
    Dim StrAddress As String
    Dim rngStart As Range
    Set rngStart = Sheet1.Range("J2")
    StrAddress = rngStart.Address

    rngStart.Cut

    'create new sheet
    Dim sht As Worksheet
    Set sht = Worksheets.Add

    'paste cell onto sheet
    Dim rngArr As Range
    Set rngArr = sht.Range("A1")
    sht.Paste rngArr

    'expand array formula size.. resize to whatever size is needed
    rngArr.Resize(3).FormulaArray = rngArr.FormulaArray

    'get your result
    Dim VarArr As Variant
    VarArr = Application.Evaluate(rngArr.CurrentArray.Address)

    ''''do something with your result here... it is an array


    'shrink the formula back to one cell
    Dim strFormula As String
    strFormula = rngArr.FormulaArray

    rngArr.CurrentArray.ClearContents
    rngArr.FormulaArray = strFormula

    'cut and paste back to original spot
    rngArr.Cut

    Sheet1.Paste Sheet1.Range(StrAddress)

    Application.DisplayAlerts = False
    sht.Delete
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

        Dim FileName As String

        FileName = path & file

        Dim wbActive As Workbook
        Set wbActive = Workbooks.Open(FileName)

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

    Dim rngInput As Range
    Set rngInput = GetInputOrSelection("Select range for waterfall")

    If rngInput Is Nothing Then
        Exit Sub
    End If

    Dim c As Range
    For Each c In Intersect(rngInput.SpecialCells(xlCellTypeBlanks), rngInput.Parent.UsedRange)
        c = c.End(xlUp)
    Next c

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

    Dim c As Range
    Set c = Range("B2")

    Dim i As Integer

    For i = 0 To 3
        c.Offset(, i) = Chr(65 + i)

        With c.Offset(1, i).Resize(10)
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

    Dim pTable As PivotTable
    Dim ws As Worksheet

    Set ws = ActiveSheet

    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."

    For Each pTable In ws.PivotTables
        Dim pField As PivotField
        For Each pField In pTable.DataFields
            On Error Resume Next
            pField.Function = xlAverage
        Next pField
    Next pTable

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

    Dim rngSelection As Range
    Set rngSelection = Application.InputBox("Select category range with heading", Type:=8)
    Set rngSelection = Intersect(rngSelection, rngSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)

    Dim rngValues As Range
    Set rngValues = Application.InputBox("Select values range with heading", Type:=8)
    Set rngValues = Intersect(rngValues, rngValues.Parent.UsedRange)

    On Error GoTo 0

    'determine default value
    Dim strDefault As Variant
    strDefault = InputBox("Enter the default value", , "#N/A")

    'detect cancel and exit
    If StrPtr(strDefault) = 0 Then
        Exit Sub
    End If

    Dim dictCategories As New Dictionary

    Dim rngCategory As Range
    For Each rngCategory In rngSelection
        'skip the header row
        If rngCategory.Address <> rngSelection.Cells(1).Address Then
            dictCategories(rngCategory.Value) = 1
        End If

    Next rngCategory

    rngValues.EntireColumn.Offset(, 1).Resize(, dictCategories.count).Insert
    'head the columns with the values

    Dim varValues As Variant
    Dim iCount As Integer
    iCount = 1
    For Each varValues In dictCategories
        rngValues.Cells(1).Offset(, iCount) = varValues
        iCount = iCount + 1
    Next varValues

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim strFormula As Variant
    strFormula = "=IF(RC" & rngSelection.Column & " =R" & _
                 rngValues.Cells(1).Row & "C,RC" & rngValues.Column & "," & strDefault & ")"

    Dim rngFormula As Range
    Set rngFormula = rngValues.Offset(1, 1).Resize(rngValues.Rows.count - 1, dictCategories.count)
    rngFormula.FormulaR1C1 = strFormula
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

    Dim rngSelection As Range
    Set rngSelection = Application.InputBox("Select category range with heading", _
                                            Type:=8)
    Set rngSelection = Intersect(rngSelection, _
                                 rngSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + _
                                                                                                xlNumbers + xlTextValues)

    Dim rngValues As Range
    Set rngValues = Application.InputBox("Select values range with heading", _
                                         Type:=8)
    Set rngValues = Intersect(rngValues, rngValues.Parent.UsedRange)

    ''need to prompt for max/min/bins
    Dim dbl_max As Double, dbl_min As Double, int_bins As Integer

    dbl_min = Application.InputBox("Minimum value.", "Min", _
                                   WorksheetFunction.Min(rngSelection), Type:=1)
    dbl_max = Application.InputBox("Maximum value.", "Max", _
                                   WorksheetFunction.Max(rngSelection), Type:=1)
    int_bins = Application.InputBox("Number of groups.", "Bins", _
                                    WorksheetFunction.RoundDown(Math.Sqr(WorksheetFunction.count(rngSelection)), _
                                                                0), Type:=1)

    On Error GoTo 0

    'determine default value
    Dim strDefault As Variant
    strDefault = Application.InputBox("Enter the default value", "Default", _
                                      "#N/A")

    'detect cancel and exit
    If StrPtr(strDefault) = 0 Then
        Exit Sub
    End If

    ''TODO prompt for output location

    rngValues.EntireColumn.Offset(, 1).Resize(, int_bins + 2).Insert
    'head the columns with the values

    ''TODO add a For loop to go through the bins

    Dim int_binNo As Integer
    For int_binNo = 0 To int_bins
        rngValues.Cells(1).Offset(, int_binNo + 1) = dbl_min + (dbl_max - _
                                                                dbl_min) * int_binNo / int_bins
    Next

    'add the last item
    rngValues.Cells(1).Offset(, int_bins + 2).FormulaR1C1 = "=RC[-1]"

    'FIRST =IF($D2 <=V$1,$U2,#N/A)
    '=IF(RC4 <=R1C,RC21,#N/A)

    'MID =IF(AND($D2 <=W$1, $D2>V$1),$U2,#N/A)  '''W current, then left
    '=IF(AND(RC4 <=R1C, RC4>R1C[-1]),RC21,#N/A)

    'LAST =IF($D2>AA$1,$U2,#N/A)
    '=IF(RC4>R1C[-1],RC21,#N/A)

    ''TODO add number format to display header correctly (helps with charts)

    'put the formula in for each column
    '=IF(RC13=R1C,RC16,#N/A)
    Dim strFormula As Variant
    strFormula = "=IF(AND(RC" & rngSelection.Column & " <=R" & _
                 rngValues.Cells(1).Row & "C," & "RC" & rngSelection.Column & ">R" & _
                 rngValues.Cells(1).Row & "C[-1]" & ")" & ",RC" & rngValues.Column & "," & _
                 strDefault & ")"

    Dim str_FirstFormula As Variant
    str_FirstFormula = "=IF(AND(RC" & rngSelection.Column & " <=R" & _
                       rngValues.Cells(1).Row & "C)" & ",RC" & rngValues.Column & "," & strDefault _
                     & ")"

    Dim str_LastFormula As Variant
    str_LastFormula = "=IF(AND(RC" & rngSelection.Column & " >R" & _
                      rngValues.Cells(1).Row & "C)" & ",RC" & rngValues.Column & "," & strDefault _
                    & ")"

    Dim rngFormula As Range
    Set rngFormula = rngValues.Offset(1, 1).Resize(rngValues.Rows.count - 1, _
                                                   int_bins + 2)
    rngFormula.FormulaR1C1 = strFormula

    'override with first/last
    rngFormula.Columns(1).FormulaR1C1 = str_FirstFormula
    rngFormula.Columns(rngFormula.Columns.count).FormulaR1C1 = str_LastFormula

    rngFormula.EntireColumn.AutoFit

    'set the number formats
    rngFormula.Offset(-1).Rows(1).Resize(1, int_bins + 1).NumberFormat = _
    "<= General"
    rngFormula.Offset(-1).Rows(1).Offset(, int_bins + 1).NumberFormat = _
    "> General"

    Exit Sub

ErrorNoSelection:
    'TODO: consider removing this prompt
    MsgBox "No selection made.  Exiting.", , "No selection"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : Sht_DeleteHiddenRows
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Deletes the hidden rows in a sheet.  Good for a "permanent" filter
'---------------------------------------------------------------------------------------
'
'Changed sub name to avoid reserved object name
Sub Sht_DeleteHiddenRows()

    Application.ScreenUpdating = False
    Dim Row As Range
    For i = ActiveSheet.UsedRange.Rows.count To 1 Step -1


        Set Row = ActiveSheet.Rows(i)

        If Row.Hidden Then
            Row.Delete
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

    Dim rng As Variant
    rng = ActiveSheet.UsedRange.Address

End Sub

