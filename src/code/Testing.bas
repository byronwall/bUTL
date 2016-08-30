Attribute VB_Name = "Testing"
Option Explicit

Sub ShowOtherOpenInstanceOfExcel()
    Dim oXLApp As Object

    'this will work if the previous instance was opened before the current one
    
    On Error Resume Next
    Set oXLApp = GetObject(, "Excel.Application")
    On Error GoTo 0

    oXLApp.Visible = True

    Set oXLApp = Nothing
End Sub

Sub PadWithSpaces()

    'quick and dirty function to add a bunch of spaces to the end of the ActiveCell

    Dim lng_spaces As Long
    lng_spaces = InputBox("How many spaces?")
    
    ActiveCell.Value = ActiveCell.Value & WorksheetFunction.Rept(" ", lng_spaces)

End Sub

Public Sub PushMergeFieldsIntoPowerPoint()

    Dim strPath As String
    strPath = "E:\Desktop\technical proposal\SulfaTrapTechnicalClarification_03182016_mcc, for merge.pptx"

    Dim appPPT As New PowerPoint.Application

    Dim presTemplate As Presentation
    Set presTemplate = appPPT.Presentations.Open(strPath)

    'create a new presentation

    Dim presNew As Presentation
    Set presNew = appPPT.Presentations.Add()

    'for each row in the table
    Dim rngData As Range
    Set rngData = Range("Table1")

    Dim rngHeaders As Range
    Set rngHeaders = Range("A1")
    Set rngHeaders = Range(rngHeaders, rngHeaders.End(xlToRight))

    presNew.ApplyTemplate (presTemplate.FullName)

    Dim rngRow As Range
    For Each rngRow In rngData.Rows
        'copy the template slide to end
        presTemplate.Slides(1).Copy
        presNew.Slides.Paste

        Dim oSld As slide
        Set oSld = presNew.Slides(presNew.Slides.count)

        Dim oShp As PowerPoint.Shape
        Dim oTxtRng As PowerPoint.TextRange
        Dim oTmpRng As PowerPoint.TextRange

        Dim rngHeader As Range
        For Each rngHeader In rngHeaders

            Dim strFind As String
            Dim strReplace As String

            strFind = rngHeader.Value
            strReplace = rngRow.Cells(1, rngHeader.Column).Value

            For Each oShp In oSld.Shapes

                If oShp.HasTextFrame Then
                    Set oTxtRng = oShp.TextFrame.TextRange

                    Set oTmpRng = oTxtRng.Replace(FindWhat:=strFind, _
                                                  Replacewhat:=strReplace, WholeWords:=False)

                    Do While Not oTmpRng Is Nothing

                        Set oTxtRng = oTxtRng.Characters(oTmpRng.start + oTmpRng.Length, _
                                                         oTxtRng.Length)

                        Set oTmpRng = oTxtRng.Replace(FindWhat:=strFind, _
                                                      Replacewhat:=strReplace, WholeWords:=False)

                    Loop
                End If

            Next oShp
        Next
        'search for each header and replace values
    Next

    presTemplate.Close

End Sub

Public Sub GetListOfMacrosCalledByButtons()
    '---------------------------------------------------------------------------------------
    ' Procedure : GetListOfMacrosCalledByButtons
    ' Author    : @byronwall
    ' Date      : 2016 01 28
    ' Purpose   : prints out a list of macros that are assigned to shapes
    '---------------------------------------------------------------------------------------
    '

    Dim sht As Worksheet
    Dim shp As Shape

    For Each sht In Worksheets
        For Each shp In sht.Shapes
            If shp.OnAction <> "" Then
                Debug.Print shp.OnAction
            End If
        Next
    Next
End Sub

Public Sub CountUnique()
    '---------------------------------------------------------------------------------------
    ' Procedure : CountUnique
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : counts the number of unique values in a Range
    '---------------------------------------------------------------------------------------
    '

    Dim rng_data As Range

    Set rng_data = GetInputOrSelection("select the range to count unique")
    Set rng_data = Intersect(rng_data, rng_data.Parent.UsedRange)

    Dim dict_vals As New Dictionary

    Dim rng_val As Range

    For Each rng_val In rng_data
        If Not dict_vals.Exists(rng_val.Value) Then
            dict_vals.Add rng_val.Value, 1
        End If
    Next

    MsgBox "items: " & dict_vals.count

End Sub

Public Sub Formula_ConcatenateCells()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_ConcatenateCells
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : will output a formula of concatenations based on cells
    '---------------------------------------------------------------------------------------
    '

    Dim rng_cell As Range
    Dim rng_joins As Range

    'get the cell to output to and the ranges to join
    Set rng_cell = GetInputOrSelection("Select the cell to put the formula")
    Set rng_joins = Application.InputBox("Select the cells to join", Type:=8)

    'get the separator
    Dim str_delim As String
    str_delim = Application.InputBox("What delimeter to use?")
    str_delim = "&""" & str_delim & """&"

    Dim arr_addr As Variant
    ReDim arr_addr(1 To rng_joins.count)

    Dim int_count As Integer
    int_count = 1

    Dim rng_join As Range
    For Each rng_join In rng_joins
        arr_addr(int_count) = rng_join.Address(False, False)
        int_count = int_count + 1
    Next

    Dim str_form As String
    str_form = "=" & Join(arr_addr, str_delim)

    rng_cell.Formula = str_form

End Sub

Public Sub Formula_ClosestInGroup()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_ClosestInGroup
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : Adds a formula that puts a given cell into a group of values based on closest value
    '---------------------------------------------------------------------------------------
    '

    Dim rng_check As Range
    Dim rng_group As Range
    Dim rng_cell As Range

    Set rng_cell = GetInputOrSelection("Select the cell to put the formula")
    Set rng_check = Application.InputBox("Select the cell to find the group of", Type:=8)
    Set rng_group = Application.InputBox("Select the group the cell belongs to", Type:=8)

    Dim str_form As String

    str_form = "=INDEX(" & rng_group.Address(True, True, xlA1, True) & _
               ",MATCH(MIN(ABS(" & rng_group.Address(True, True, xlA1, True) & "-" & _
               rng_check.Address(False, False) & ")),ABS(" & rng_group.Address(True, True, xlA1, True) & "-" & rng_check.Address(False, False) & "),0))"

    rng_cell.FormulaArray = str_form

End Sub

Public Sub SelectAllArrayFormulas()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectAllArrayFormulas
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : selects all cells on current sheet that have an array formula
    '---------------------------------------------------------------------------------------
    '

    Dim rng_forms As Range

    Set rng_forms = ActiveSheet.UsedRange

    Dim rng_select As Range

    Dim rng_form As Range
    For Each rng_form In rng_forms
        If rng_form.HasArray Then
            If rng_select Is Nothing Then
                Set rng_select = rng_form
            Else
                Set rng_select = Union(rng_select, rng_form)
            End If
        End If
    Next

    rng_select.Select

End Sub

Public Sub CharacterCodesForSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : CharacterCodesForSelection
    ' Author    : @byronwall
    ' Date      : 2016 01 27
    ' Purpose   : will output each character in the text
    '---------------------------------------------------------------------------------------
    '

    Dim letter As Variant

    Dim rng_val As Range
    Set rng_val = Selection

    Dim i As Integer
    For i = 1 To Len(rng_val.Value)
        MsgBox Asc(Mid(rng_val.Value, i, 1))
    Next

End Sub

Public Sub Formula_CreateCountNameForArray()
    '---------------------------------------------------------------------------------------
    ' Procedure : Formula_CreateCountNameForArray
    ' Author    : @byronwall
    ' Date      : 2016 01 21
    ' Purpose   : meant to create formula with limited range of column
    '---------------------------------------------------------------------------------------
    '

    Dim rng_named As Range

    Dim str_name As String
    str_name = Application.InputBox("Name of the range", Type:=2)

    Set rng_named = ActiveWorkbook.Names(str_name).RefersToRange

    Dim str_form As String
    str_form = "=INDEX(" & str_name & ",1,1):INDEX(" & str_name & ",COUNTA(" & str_name & "),1)"

    ActiveWorkbook.Names.Add str_name & "_limited", str_form

End Sub

Public Sub CopyDiscontinuousRangeValuesToClipboard()

    Dim rngCSV As Range
    Set rngCSV = GetInputOrSelection("Choose range for converting to CSV")

    If rngCSV Is Nothing Then
        Exit Sub
    End If

    'get the counts for rows/columns
    Dim int_row As Integer
    Dim int_cols As Integer

    Set rngCSV = Intersect(rngCSV, rngCSV.Parent.UsedRange)

    'build the string array
    Dim arr_rows() As String
    ReDim arr_rows(1 To rngCSV.Areas(1).Rows.count) As String

    Dim bool_firstArea As Boolean
    bool_firstArea = True

    Dim rng_area As Range
    For Each rng_area In rngCSV.Areas
        For int_row = 1 To UBound(arr_rows)
            If bool_firstArea Then
                arr_rows(int_row) = Join(Application.Transpose(Application.Transpose(rng_area.Rows(int_row).Value)), vbTab)
            Else
                arr_rows(int_row) = arr_rows(int_row) & vbTab & Join(Application.Transpose(Application.Transpose(rng_area.Rows(int_row).Value)), vbTab)
            End If
        Next

        bool_firstArea = False
    Next

    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    clipboard.SetText Join(arr_rows, vbCrLf)
    clipboard.PutInClipboard

End Sub

Public Sub ComputeDistanceMatrix()

    'get the range of inputs, along with input name
    Dim rng_input As Range
    Set rng_input = Application.InputBox("Select input data", "Input", Type:=8)

    'Dim rng_ID As Range
    'Set rng_ID = Application.InputBox("Select ID data", "ID", Type:=8)

    'turning off updates makes a huge difference here... could also use array for output
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'create new workbook
    Dim wkbk As Workbook
    Set wkbk = Workbooks.Add

    Dim sht_out As Worksheet
    Set sht_out = wkbk.Sheets(1)
    sht_out.name = "scaled data"

    'copy data over to standardize
    rng_input.Copy wkbk.Sheets(1).Range("A1")

    'go to edge of data, add a column, add STANDARDIZE, copy paste values, delete

    Dim rng_data As Range
    Set rng_data = sht_out.Range("A1").CurrentRegion

    Dim rng_col As Range
    For Each rng_col In rng_data.Columns

        'edge cell
        Dim rng_edge As Range
        Set rng_edge = sht_out.Cells(1, sht_out.Columns.count).End(xlToLeft).Offset(, 1)

        'do a normal dist standardization
        '=STANDARDIZE(A1,AVERAGE(A:A),STDEV.S(A:A))

        rng_edge.Formula = "=IFERROR(STANDARDIZE(" & rng_col.Cells(1, 1).Address(False, False) & ",AVERAGE(" & _
                           rng_col.Address & "),STDEV.S(" & rng_col.Address & ")),0)"

        'do a simple value over average to detect differences
        rng_edge.Formula = "=IFERROR(" & rng_col.Cells(1, 1).Address(False, False) & "/AVERAGE(" & _
                           rng_col.Address & "),1)"

        'fill that down
        Range(rng_edge, rng_edge.Offset(, -1).End(xlDown).Offset(, 1)).FillDown

    Next

    Application.Calculate
    sht_out.UsedRange.Value = sht_out.UsedRange.Value
    rng_data.EntireColumn.Delete

    Dim sht_dist As Worksheet
    Set sht_dist = wkbk.Worksheets.Add()
    sht_dist.name = "distances"

    Dim rng_out As Range
    Set rng_out = sht_dist.Range("A1")

    'loop through each row with each other row
    Dim rng_row1 As Range
    Dim rng_row2 As Range

    Set rng_input = sht_out.Range("A1").CurrentRegion

    For Each rng_row1 In rng_input.Rows
        For Each rng_row2 In rng_input.Rows

            'loop through each column and compute the distance
            Dim dbl_dist_sq As Double
            dbl_dist_sq = 0

            Dim int_col As Integer
            For int_col = 1 To rng_row1.Cells.count
                dbl_dist_sq = dbl_dist_sq + (rng_row1.Cells(1, int_col) - rng_row2.Cells(1, int_col)) ^ 2
            Next

            'take the sqrt of that value and output
            rng_out.Value = dbl_dist_sq ^ 0.5

            'get to next column for output
            Set rng_out = rng_out.Offset(, 1)
        Next

        'drop down a row and go back to left edge
        Set rng_out = rng_out.Offset(1).End(xlToLeft)
    Next

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    sht_dist.UsedRange.NumberFormat = "0.00"
    sht_dist.UsedRange.EntireColumn.AutoFit

    'do the coloring
    Formatting_AddCondFormat sht_dist.UsedRange

End Sub

Sub RemoveAllLegends()

    Dim cht_obj As ChartObject
    
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        cht_obj.Chart.HasLegend = False
        cht_obj.Chart.HasTitle = True
        
        cht_obj.Chart.SeriesCollection(1).MarkerSize = 4
    Next

End Sub

Sub ApplyFormattingToEachColumn()
    Dim rng As Range
    For Each rng In Selection.Columns

        Formatting_AddCondFormat rng
    Next
End Sub

Private Sub Formatting_AddCondFormat(ByVal rng As Range)

        rng.FormatConditions.AddColorScale ColorScaleType:=3
        rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
        rng.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
        With rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 7039480
            .TintAndShade = 0
        End With
        rng.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
        rng.FormatConditions(1).ColorScaleCriteria(2).Value = 50
        With rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 8711167
            .TintAndShade = 0
        End With
        rng.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 8109667
            .TintAndShade = 0
        End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : TraceDependentsForAll
' Author    : @byronwall
' Date      : 2015 11 09
' Purpose   : Quick Sub to iterate through Selection and Trace Dependents for all
'---------------------------------------------------------------------------------------
'
Sub TraceDependentsForAll()

    Dim rng As Range
    
    For Each rng In Intersect(Selection, Selection.Parent.UsedRange)
        rng.ShowDependents
    Next rng

End Sub

