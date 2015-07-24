Attribute VB_Name = "Usability"
Function GetRow(arr As Variant, ResultArr As Variant, RowNumber As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetRow
' This populates ResultArr with a one-dimensional array that is the
' specified row of Arr. The existing contents of ResultArr are
' destroyed. ResultArr must be a dynamic array.
' Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ColNdx As Long

''''''''''''''''''''''''''''''''''''
' Ensure ColumnNumber is less than
' or equal to the number of columns.
''''''''''''''''''''''''''''''''''''
If UBound(arr, 1) < RowNumber Then
    GetRow = False
    Exit Function
End If
If LBound(arr, 1) > RowNumber Then
    GetRow = False
    Exit Function
End If

Erase ResultArr
ReDim ResultArr(LBound(arr, 2) To UBound(arr, 2))
For ColNdx = LBound(ResultArr) To UBound(ResultArr)
    ResultArr(ColNdx) = arr(RowNumber, ColNdx)
Next ColNdx

GetRow = True


End Function




Sub ForceRecalc()

    Application.CalculateFullRebuild

End Sub

Sub CombineAllSheetsData()

    '''code will combine all the data, reusing columns where overlapping
    
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
                wsData.rows(1).Copy wsCombined.Range("A1")
            
                boolFirst = False
            Else
                'search for missing columns
                Dim rngHeader As Range
                For Each rngHeader In Intersect(wsData.rows(1), wsData.UsedRange)
                    
                    'check if it exists
                    Dim varHdrMatch As Variant
                    varHdrMatch = Application.Match(rngHeader, wsCombined.rows(1), 0)
                    
                    'if not, add to header row
                    If IsError(varHdrMatch) Then
                        wsCombined.Range("A1").End(xlToRight).Offset(, 1) = rngHeader
                    End If
                Next rngHeader
            End If
            
            'find the PnPID column for combo
            Dim int_colId As Integer
            int_colId = Application.Match("PnPID", wsCombined.rows(1), 0)
            
            'find the PnPID column for data
            Dim iColIDData As Integer
            iColIDData = Application.Match("PnPID", wsData.rows(1), 0)
            
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
                        iDataRow = wsCombined.Columns(int_colId).Cells(wsCombined.rows.count, 1).End(xlUp).Offset(1).Row
                        wsCombined.Cells(iDataRow, int_colId) = wsData.Cells(c.Row, iColIDData)
                    End If
                    
                    'get column
                    Dim iCol As Integer
                    iCol = Application.Match(wsData.Cells(1, c.Column), wsCombined.rows(1), 0)
                    
                    'update combo data
                    wsCombined.Cells(iDataRow, iCol) = c
                    
                End If
            Next c
        End If
    Next wsData
End Sub

Sub ConvertSelectionToCsv()

    'this will go through a block of data and return a single CSV string
    
    Dim csvRow As Range
    
    Dim rngCSV As Range
    
    Set rngCSV = Selection
    
    Dim csvOut As String
    csvOut = ""
    
    
    For Each csvRow In rngCSV.rows
        Dim arr() As Variant
        GetRow csvRow.rows.Value, arr, 1
        
        csvOut = csvOut & Join(arr, ",") & vbCrLf
        
    Next csvRow
    
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    
    clipboard.SetText csvOut
    clipboard.PutInClipboard

End Sub

Sub CopyClear()

    'Save the selection
    Dim rngSource As Range
    Set rngSource = Selection
    
    'Determine the destination
    Dim rngDestination As Range
    Set rngDestination = Application.InputBox("Select the destination", Type:=8)
    
    'Freeze screen
    Application.ScreenUpdating = False
    
    'Copy the source
    rngSource.Copy
    
    'Determine the offset of change
    Dim rngOff As Range
    Set rngOff = rngSource.Offset(rngDestination.Row - rngSource.Row, rngDestination.Column - rngSource.Column)
    
    'Paste to the destination
    rngDestination.PasteSpecial xlPasteAll
    
    'Clear any cells that were in the source and not in the destination
    Dim rngClear As Range
    For Each rngClear In rngSource
        If Intersect(rngClear, rngOff) Is Nothing Then
            rngClear.Clear
        End If
    Next rngClear
 End Sub
Sub ColorInputs()

    Dim c As Range
    
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


Sub CutPasteTranspose()


    Dim rngSelect As Range
    Set rngSelect = rngSelectection
    
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
    
    Dim c As Range
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
End Sub

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

'''"TEST
Sub ExportFilesFromFolder()

    Application.ScreenUpdating = False

    Dim file As Variant
    ' This path is constant?
    Dim path As String
    path = "C:\Users\eltron\Desktop\PSV Sizing\Completed\"
    
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

'Sub will take a selection of values and fill blanks with the value from above
'#BUTTON
Sub FillValueDown()

    Dim rngInput As Range
    Set rngInput = Selection
    
    Dim c As Range
    
    For Each c In Intersect(rngInput.SpecialCells(xlCellTypeBlanks), Selection.Parent.UsedRange)
        c = c.End(xlUp)
    Next c

End Sub

Sub GenerateRandomData()

    '''code will create 1 column of Dates followed by three columns of integers
    'Dates will be sequential, Headers in Row 2 beginning in Column B with 10 data points below
    
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

Sub OpenContainingFolder()

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    If wb.path <> "" Then
        wb.FollowHyperlink wb.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub

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

'this code is used to break a value out into categories
Sub SeriesSplit()

    'find the unique values in the category field (assumes header and entire column)
    Dim rngSelection As Range
    Set rngSelection = Application.InputBox("Select category range with heading", Type:=8)
    Set rngSelection = Intersect(rngSelection, rngSelection.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)
    
    Dim dictCategories As New Dictionary
    
    Dim rngCategory As Range
    For Each rngCategory In rngSelection
        'skip the header row
        If rngCategory.Address = rngSelection.Cells(1).Address Then
        
        Else
            dictCategories(rngCategory.Value) = 1
        End If
        
    Next rngCategory
    
    'create that number of columns next to value column
    
    Dim rngValues As Range
    Set rngValues = Application.InputBox("Select values range with heading", Type:=8)
    Set rngValues = Intersect(rngValues, rngValues.Parent.UsedRange)
        
    rngValues.EntireColumn.Offset(, 1).Resize(, dictCategories.count).Insert
    'head the columns with the values
    
    Dim varValues As Variant
    Dim iCount As Integer
    iCount = 1
    For Each varValues In dictCategories
        rngValues.Cells(1).Offset(, iCount) = varValues
        iCount = iCount + 1
    Next varValues
    
    'determine default value
    Dim strDefault As Variant
    strDefault = InputBox("Enter the default value", , """""")
    
    'put the formula in for each column
    'FORMULA
    '=IF(RC13=R1C,RC16,#N/A)
    Dim strFormula As Variant
    strFormula = "=IF(RC" & rngSelection.Column & " =R" & rngValues.Cells(1).Row & "C,RC" & rngValues.Column & "," & strDefault & ")"
    
    Dim rngFormula As Range
    Set rngFormula = rngValues.Offset(1, 1).Resize(rngValues.rows.count - 1, dictCategories.count)
    rngFormula.FormulaR1C1 = strFormula
    rngFormula.EntireColumn.AutoFit
End Sub

Sub Sheet_DeleteHiddenRows()

    'this sub will delete all the hidden rows
    'this would be used with a filter to pare down the list
    
    Application.ScreenUpdating = False
    Dim Row As Range
    For i = ActiveSheet.UsedRange.rows.count To 1 Step -1
        
        
        Set Row = ActiveSheet.rows(i)
        
        If Row.Hidden Then
            Row.Delete
        End If
    Next i

    Application.ScreenUpdating = True

End Sub

Sub UnhideAllRowsAndColumns()

    'need to unhide rows and then columns
    
    ActiveSheet.Cells.EntireRow.Hidden = False
    ActiveSheet.Cells.EntireColumn.Hidden = False

End Sub

Sub UpdateScrollbars()
    'What does this do?
    Dim rng As Variant
    rng = ActiveSheet.UsedRange.Address

End Sub












