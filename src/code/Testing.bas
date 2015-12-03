Attribute VB_Name = "Testing"
Option Explicit

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

