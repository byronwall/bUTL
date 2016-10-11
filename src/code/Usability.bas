Attribute VB_Name = "Usability"
Option Explicit

Sub ColorInputs()
    '---------------------------------------------------------------------------------------
    ' Procedure : ColorInputs
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Finds cells with no value and colors them based on having a formula?
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
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


Sub CombineAllSheetsData()
    '---------------------------------------------------------------------------------------
    ' Procedure : CombineAllSheetsData
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Combines all sheets, resuing columns where the same
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'create the new wkbk and sheet
    Dim wbCombo As Workbook
    Dim wbData As Workbook

    Set wbData = ActiveWorkbook
    Set wbCombo = Workbooks.Add

    Dim wsCombined As Worksheet
    Set wsCombined = wbCombo.Sheets.Add

    Dim boolFirst As Boolean
    boolFirst = True

    Dim iComboRow As Long
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
            Dim int_colId As Long
            int_colId = Application.Match("PnPID", wsCombined.Rows(1), 0)

            'find the PnPID column for data
            Dim iColIDData As Long
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
                    Dim iCol As Long
                    iCol = Application.Match(wsData.Cells(1, c.Column), wsCombined.Rows(1), 0)

                    'update combo data
                    wsCombined.Cells(iDataRow, iCol) = c

                End If
            Next c
        End If
    Next wsData
End Sub


Sub ConvertSelectionToCsv()
    '---------------------------------------------------------------------------------------
    ' Procedure : ConvertSelectionToCsv
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Crude CSV output from the current selection, works with numbers
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim rngCSV As Range
    Set rngCSV = GetInputOrSelection("Choose range for converting to CSV")

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

Public Sub CopyCellAddress()
    '---------------------------------------------------------------------------------------
    ' Procedure : CopyCellAddress
    ' Author    : @byronwall
    ' Date      : 2015 12 03
    ' Purpose   : Copies the current cell address to the clipboard for paste use in a formula
    '---------------------------------------------------------------------------------------
    '

    'TODO: this need to get a button or a keyboard shortcut for easy use
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    Dim rng_sel As Range
    Set rng_sel = Selection

    clipboard.SetText rng_sel.Address(True, True, xlA1, True)
    clipboard.PutInClipboard
End Sub

Sub Sheet_DeleteHiddenRows()
    'These rows are unrecoverable
    Dim shouldDeleteHiddenRows As VbMsgBoxResult
    shouldDeleteHiddenRows = MsgBox("This will permanently delete hidden rows. They cannot be recovered. Are you sure?", vbYesNo)
    
    If Not shouldDeleteHiddenRows = vbYes Then
        Exit Sub
    End If
        
    Application.ScreenUpdating = False
    
    'collect a range to delete at end, using UNION-DELETE
    Dim rngToDelete As Range
    
    Dim iCount As Long
    iCount = 0
    With ActiveSheet
        Dim rowIndex As Long
        For rowIndex = .UsedRange.Rows.count To 1 Step -1
            If .Rows(rowIndex).Hidden Then
                If rngToDelete Is Nothing Then
                    Set rngToDelete = .Rows(rowIndex)
                Else
                    Set rngToDelete = Union(rngToDelete, .Rows(rowIndex))
                End If
                iCount = iCount + 1
            End If
        Next rowIndex
    End With
    
    rngToDelete.Delete
    
    Application.ScreenUpdating = True
    
    MsgBox (iCount & " rows were deleted")
End Sub


Sub CutPasteTranspose()
    '---------------------------------------------------------------------------------------
    ' Procedure : CutPasteTranspose
    ' Author    : @byronwall, @RaymondWise
    ' Date      : 2015 07 31
    ' Purpose   : Does a cut/transpose by cutting each cell individually
    '---------------------------------------------------------------------------------------
    '

    '########Still Needs to address Issue#23#############
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

    Dim iCRow As Long
    iCRow = rCorner.Row
    Dim iCCol As Long
    iCCol = rCorner.Column

    Dim iORow As Long
    Dim iOCol As Long
    iORow = rngOut.Row
    iOCol = rngOut.Column

    rngOut.Activate
    
    'Check to not overwrite
    Dim c As Range
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

Sub FillValueDown()
    '---------------------------------------------------------------------------------------
    ' Procedure : FillValueDown
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Does a fill of blank values from the cell above with a value
    '---------------------------------------------------------------------------------------
    '
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


Sub ForceRecalc()
    '---------------------------------------------------------------------------------------
    ' Procedure : ForceRecalc
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Provides a button to do a full recalc
    '---------------------------------------------------------------------------------------
    '
    Application.CalculateFullRebuild

End Sub


Sub GenerateRandomData()
    '---------------------------------------------------------------------------------------
    ' Procedure : GenerateRandomData
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Generates a block of random data for testing questions on SO
    '---------------------------------------------------------------------------------------
    '
    Dim c As Range
    Set c = Range("B2")

    Dim i As Long

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


Sub OpenContainingFolder()
    '---------------------------------------------------------------------------------------
    ' Procedure : OpenContainingFolder
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Open the folder that contains the ActiveWorkbook
    '---------------------------------------------------------------------------------------
    '
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    If wb.path <> "" Then
        wb.FollowHyperlink wb.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub


Sub PivotSetAllFields()
    '---------------------------------------------------------------------------------------
    ' Procedure : PivotSetAllFields
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets all fields in a PivotTable to use a certain calculation type
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
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

Sub SeriesSplit()
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesSplit
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Takes a category columns and splits the values out into new columns for each unique entry
    '---------------------------------------------------------------------------------------
    '
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
    Dim iCount As Long
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


Sub SeriesSplitIntoBins()
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesSplitIntoBins
    ' Author    : @byronwall
    ' Date      : 2015 11 03
    ' Purpose   : Code will break a column of continuous data into bins for plotting
    '---------------------------------------------------------------------------------------
    '
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
    Dim dbl_max As Double, dbl_min As Double, int_bins As Long

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

    Dim int_binNo As Long
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


Sub UnhideAllRowsAndColumns()
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


Sub UpdateScrollbars()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateScrollbars
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Cheap trick that forces Excel to update the scroll bars after a large deletion
    '---------------------------------------------------------------------------------------
    '
    Dim rng As Variant
    rng = ActiveSheet.UsedRange.Address

End Sub

