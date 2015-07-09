Attribute VB_Name = "Usability"
Sub EvaluateArrayFormulaOnNewSheet()

    'cut cell with formula
    Dim str_address As String
    Dim rng_start As Range
    Set rng_start = Sheet1.Range("J2")
    str_address = rng_start.Address
    
    rng_start.Cut
    
    'create new sheet
    Dim sht As Worksheet
    Set sht = Worksheets.Add
    
    'paste cell onto sheet
    Dim rng_arr As Range
    Set rng_arr = sht.Range("A1")
    sht.Paste rng_arr
    
    'expand array formula size.. resize to whatever size is needed
    rng_arr.Resize(3).FormulaArray = rng_arr.FormulaArray
    
    'get your result
    Dim v_arr As Variant
    v_arr = Application.Evaluate(rng_arr.CurrentArray.Address)
    
    ''''do something with your result here... it is an array
    
    
    'shrink the formula back to one cell
    Dim str_formula As String
    str_formula = rng_arr.FormulaArray
    
    rng_arr.CurrentArray.ClearContents
    rng_arr.FormulaArray = str_formula
    
    'cut and paste back to original spot
    rng_arr.Cut
    
    Sheet1.Paste Sheet1.Range(str_address)
    
    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True

End Sub

Sub GenerateRandomData()

    '''code will create 4 columns of random data of different types
    
    Dim rng_cell As Range
    Set rng_cell = Range("B2")
    
    Dim i As Integer
    
    For i = 0 To 3
        rng_cell.Offset(, i) = Chr(65 + i)
        
        With rng_cell.Offset(1, i).Resize(10)
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

Sub CombineAllSheetsData()

    '''code will combine all the data, reusing columns where overlapping
    
    'create the new wkbk and sheet
    Dim wkbk_combo As Workbook
    Dim wkbk_data As Workbook
    
    Set wkbk_data = ActiveWorkbook
    Set wkbk_combo = Workbooks.Add
    
    Dim sht_combined As Worksheet
    Set sht_combined = wkbk_combo.Sheets.Add
    
    Dim bool_first As Boolean
    bool_first = True
    
    Dim int_comboRow As Integer
    int_comboRow = 1
    
    Dim sht_data As Worksheet
    For Each sht_data In wkbk_data.Sheets
        If sht_data.name <> sht_combined.name Then
        
            sht_data.Unprotect
            
            'get the headers squared up
            If bool_first Then
                'copy over all headers
                sht_data.rows(1).Copy sht_combined.Range("A1")
            
                bool_first = False
            Else
                'search for missing columns
                Dim rng_header As Range
                For Each rng_header In Intersect(sht_data.rows(1), sht_data.UsedRange)
                    
                    'check if it exists
                    Dim hdr_match As Variant
                    hdr_match = Application.Match(rng_header, sht_combined.rows(1), 0)
                    
                    'if not, add to header row
                    If IsError(hdr_match) Then
                        sht_combined.Range("A1").End(xlToRight).Offset(, 1) = rng_header
                    End If
                Next rng_header
            End If
            
            'find the PnPID column for combo
            Dim int_colId As Integer
            int_colId = Application.Match("PnPID", sht_combined.rows(1), 0)
            
            'find the PnPID column for data
            Dim int_colIdData As Integer
            int_colIdData = Application.Match("PnPID", sht_data.rows(1), 0)
            
            'add the data, row by row
            Dim rng_cell As Range
            For Each rng_cell In sht_data.UsedRange.SpecialCells(xlCellTypeConstants)
                If rng_cell.row > 1 Then
                
                    'check if the PnPID exists in the combo sheet
                    Dim int_dataRow As Variant
                    int_dataRow = Application.Match( _
                        sht_data.Cells(rng_cell.row, int_colIdData), _
                        sht_combined.Columns(int_colId), _
                        0)
                        
                    'add new row if it did not exist and id number
                    If IsError(int_dataRow) Then
                        int_dataRow = sht_combined.Columns(int_colId).Cells(sht_combined.rows.count, 1).End(xlUp).Offset(1).row
                        sht_combined.Cells(int_dataRow, int_colId) = sht_data.Cells(rng_cell.row, int_colIdData)
                    End If
                    
                    'get column
                    Dim int_column As Integer
                    int_column = Application.Match(sht_data.Cells(1, rng_cell.Column), sht_combined.rows(1), 0)
                    
                    'update combo data
                    sht_combined.Cells(int_dataRow, int_column) = rng_cell
                    
                End If
            Next rng_cell
        End If
    Next sht_data
End Sub

'''"TEST
Sub ExportFilesFromFolder()

    Application.ScreenUpdating = False

    Dim file As Variant
    
    Dim str_dir As String
    str_dir = "C:\Users\eltron\Desktop\PSV Sizing\Completed\"
    
   file = Dir(str_dir)
   While (file <> "")
   
    Debug.Print str_dir & file
    
    Dim str_filename As String
    
    str_filename = str_dir & file
    
    Dim wkbk As Workbook
    Set wkbk = Workbooks.Open(str_filename)
    
    Dim wksht As Worksheet
    Set wksht = wkbk.Sheets("Case Summary")
    
     With ActiveSheet.PageSetup
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
    End With

    wksht.ExportAsFixedFormat xlTypePDF, str_dir & "PDFs\" & file & ".pdf"
    
    wkbk.Close False
   
     file = Dir
  Wend
  
  Application.ScreenUpdating = True

End Sub


Sub UnhideAllRowsAndColumns()

    'need to unhide rows and then columns
    
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    
    sht.Cells.EntireRow.Hidden = False
    sht.Cells.EntireColumn.Hidden = False

End Sub

Sub ColorInputs()

    Dim cell As Range
    
    For Each cell In Selection
        If cell.Value <> "" Then
            If cell.HasFormula Then
                cell.Interior.ThemeColor = msoThemeColorAccent1
            Else
                cell.Interior.ThemeColor = msoThemeColorAccent2
            End If
        End If
    Next cell

End Sub

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
Sub ConvertSelectionToCsv()

    'this will go through a block of data and return a single CSV string
    Dim cell As Range
    Dim csv_row As Range
    Dim csv_col As Range
    
    Dim csv_all As Range
    
    Set csv_all = Selection
    
    Dim csv_output As String
    csv_output = ""
    
    
    For Each csv_row In csv_all.rows
        Dim arr() As Variant
        GetRow csv_row.rows.Value, arr, 1
        
        csv_output = csv_output & Join(arr, ",") & vbCrLf
        
    Next csv_row
    
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    
    clipboard.SetText csv_output
    clipboard.PutInClipboard

End Sub

Sub Sheet_DeleteHiddenRows()

    'this sub will delete all the hidden rows
    'this would be used with a filter to pare down the list
    
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    
    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    For i = sht.UsedRange.rows.count To 1 Step -1
        Dim rng_row As Range
        
        Set rng_row = sht.rows(i)
        
        If rng_row.Hidden Then
            rng_row.Delete
        End If
    Next i

    Application.ScreenUpdating = True

End Sub

Sub PivotSetAllFields()

    Dim pivot As PivotTable
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    
    MsgBox "This defaults to the average for every Pivot table on the sheet.  Edit code for other result."
    
    For Each pivot In sht.PivotTables
        Dim pivot_field As PivotField
        For Each pivot_field In pivot.DataFields
            On Error Resume Next
            pivot_field.Function = xlAverage
        Next pivot_field
    Next pivot

End Sub

'this code is used to break a value out into categories
Sub SeriesSplit()

    'find the unique values in the category field (assumes header and entire column)
    Dim rng_cats As Range
    Set rng_cats = Application.InputBox("Select category range with heading", Type:=8)
    Set rng_cats = Intersect(rng_cats, rng_cats.Parent.UsedRange).SpecialCells(xlCellTypeVisible, xlLogical + xlNumbers + xlTextValues)
    
    Dim dict_cat As New Dictionary
    
    Dim rng_cat As Range
    For Each rng_cat In rng_cats
        'skip the header row
        If rng_cat.Address = rng_cats.Cells(1).Address Then
        
        Else
            dict_cat(rng_cat.Value) = 1
        End If
        
    Next rng_cat
    
    'create that number of columns next to value column
    
    Dim rng_values As Range
    Set rng_values = Application.InputBox("Select values range with heading", Type:=8)
    Set rng_values = Intersect(rng_values, rng_values.Parent.UsedRange)
        
    rng_values.EntireColumn.Offset(, 1).Resize(, dict_cat.count).Insert
    'head the columns with the values
    
    Dim var_value As Variant
    Dim int_count As Integer
    int_count = 1
    For Each var_value In dict_cat
        rng_values.Cells(1).Offset(, int_count) = var_value
        int_count = int_count + 1
    Next var_value
    
    'determine default value
    Dim str_default As Variant
    str_default = InputBox("Enter the default value", , """""")
    
    'put the formula in for each column
    'FORMULA
    '=IF(RC13=R1C,RC16,#N/A)
    Dim str_formula As Variant
    str_formula = "=IF(RC" & rng_cats.Column & " =R" & rng_values.Cells(1).row & "C,RC" & rng_values.Column & "," & str_default & ")"
    
    Dim rng_formula As Range
    Set rng_formula = rng_values.Offset(1, 1).Resize(rng_values.rows.count - 1, dict_cat.count)
    rng_formula.FormulaR1C1 = str_formula
    rng_formula.EntireColumn.AutoFit

End Sub

Sub ForceRecalc()

    Application.CalculateFullRebuild

End Sub

Sub CopyClear()

    'Save the selection
    Dim rng_src As Range
    Set rng_src = Selection
    
    'Determine the destination
    Dim cell_dest As Range
    Set cell_dest = Application.InputBox("Select the destination", Type:=8)
    
    'Freeze screen
    Application.ScreenUpdating = False
    
    'Copy the source
    rng_src.Copy
    
    'Determine the offset of change
    Dim rng_dest As Range
    Set rng_dest = rng_src.Offset(cell_dest.row - rng_src.row, cell_dest.Column - rng_src.Column)
    
    'Paste to the destination
    cell_dest.PasteSpecial xlPasteAll
    
    'Clear any cells that were in the source and not in the destination
    Dim cell_clr As Range
    For Each cell_clr In rng_src
        If Intersect(cell_clr, rng_dest) Is Nothing Then
            cell_clr.Clear
        End If
    Next cell_clr
    

End Sub

Sub UpdateScrollbars()
    
    Dim rng As Variant
    rng = ActiveSheet.UsedRange.Address

End Sub

Sub OpenContainingFolder()

    Dim wkbk As Workbook
    Set wkbk = ActiveWorkbook
    
    If wkbk.path <> "" Then
        wkbk.FollowHyperlink wkbk.path
    Else
        MsgBox "Open file is not in a folder yet."
    End If

End Sub

Sub CutPasteTranspose()


    Dim sel As Range
    Set sel = Selection
    
    Dim rng_out As Range
    Set rng_out = Application.InputBox("Select output corner", Type:=8)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual


    
    Dim rng_corner As Range
    Set rng_corner = sel.Cells(1, 1)
    
    Dim corn_row, corn_col
    corn_row = rng_corner.row
    corn_col = rng_corner.Column
    
    Dim rng_row, rng_col
    rng_row = rng_out.row
    rng_col = rng_out.Column
    
    rng_out.Activate
    
    Dim cell As Range
    For Each cell In sel
        cell.Cut
        ActiveSheet.Cells(rng_row + cell.Column - corn_col, rng_col + cell.row - corn_row).Activate
        ActiveSheet.Paste
    Next cell
    
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

End Sub



'Sub will take a selection of values and fill blanks with the value from above
'#BUTTON
Sub FillValueDown()

    Dim rng_in As Range
    Set rng_in = Selection
    
    Dim rng_cell As Range
    
    For Each rng_cell In Intersect(rng_in.SpecialCells(xlCellTypeBlanks), Selection.Parent.UsedRange)
        rng_cell = rng_cell.End(xlUp)
    Next rng_cell

End Sub


