Attribute VB_Name = "Testing"
Option Explicit

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

    Dim rng_out As Range
    Set rng_out = wkbk.Sheets(1).Range("A1")

    'loop through each row with each other row
    Dim rng_row1 As Range
    Dim rng_row2 As Range

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

End Sub

Sub RemoveAllLegends()

    Dim cht_obj As ChartObject
    
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        cht_obj.Chart.HasLegend = False
        cht_obj.Chart.HasTitle = False
    Next

End Sub

Sub ApplyFormattingToEachColumn()
'
' Macro1 Macro
'

'
    Dim rng As Range
    For Each rng In Selection.Columns

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
    Next
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

