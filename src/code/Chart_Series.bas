Attribute VB_Name = "Chart_Series"
Option Explicit

Sub Chart_AddTrendlineToSeriesAndColor()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AddTrendlineToSeriesAndColor
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds a trendline to each series in all charts
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        Dim chartIndex As Long
        chartIndex = 1
        
        Dim chtSeries As series
        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries chtSeries

            'clear out old ones
            Dim j As Long
            For j = 1 To chtSeries.Trendlines.count
                chtSeries.Trendlines(j).Delete
            Next j

            chtSeries.MarkerBackgroundColor = Chart_GetColor(chartIndex)

            Dim trend As Trendline
            Set trend = chtSeries.Trendlines.Add()
            trend.Type = xlLinear
            trend.Border.Color = chtSeries.MarkerBackgroundColor
            
            '2015 11 06 test to avoid error without name
            '2015 12 07 dealing with multi-cell Names
            'TODO: handle if the name is not a range also
            If Not butlSeries.name Is Nothing Then
                trend.name = butlSeries.name.Cells(1, 1).Value
            End If

            trend.DisplayEquation = True
            trend.DisplayRSquared = True
            trend.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(chartIndex)

            chartIndex = chartIndex + 1
        Next chtSeries

    Next chtObj
End Sub


Sub Chart_ExtendSeriesToRanges()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ExtendSeriesToRanges
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Extends the underlying data for a series to go to the end of its current Range
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series

        'get each series
        For Each chtSeries In chtObj.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries chtSeries

            If Not butlSeries.XValues Is Nothing Then
                chtSeries.XValues = RangeEnd(butlSeries.XValues.Cells(1), xlDown)
            End If
            chtSeries.Values = RangeEnd(butlSeries.Values.Cells(1), xlDown)

        Next chtSeries
    Next chtObj
End Sub


Sub Chart_GoToXRange()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GoToXRange
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Selects the x value range that is used for the series
    '---------------------------------------------------------------------------------------
    '

    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.XValues.Parent.Activate
        b.XValues.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub


Sub Chart_GoToYRange()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GoToYRange
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Selects the y values used for the series
    '---------------------------------------------------------------------------------------
    '

    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.Values.Parent.Activate
        b.Values.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub


Sub Chart_RemoveTrendlines()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_RemoveTrendlines
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Remove all trendlines from a chart
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series
        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim trend As Trendline
            For Each trend In chtSeries.Trendlines
                trend.Delete
            Next trend
        Next chtSeries
    Next chtObj
End Sub


Sub Chart_RerangeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_RerangeSeries
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Entry point for an interface to help rerange series
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim frm As New form_chtSeries
    frm.Show

End Sub


Sub Chart_TrendlinesToAverage()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TrendlinesToAverage
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a trendline using a moving average instead of linear
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series

        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In chtSeries.Trendlines
                trend.Type = xlMovingAvg
                trend.Period = 15
                trend.Format.Line.Weight = 2
            Next
        Next
    Next

End Sub


Sub ChartFlipXYValues()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartFlipXYValues
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Flips the x/y ranges for each series
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    Dim cht As Chart
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        Set cht = chtObj.Chart

        Dim butlSeriesies As New Collection
        Dim butlSeries As bUTLChartSeries
        
        Dim chtSeries As series
        For Each chtSeries In cht.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries chtSeries

            Dim rngDummy As Range

            Set rngDummy = butlSeries.Values
            Set butlSeries.Values = butlSeries.XValues
            Set butlSeries.XValues = rngDummy

            'need to change the series name also
            'assume that title is same offset
            'code blocked for now
            If False And Not butlSeries.name Is Nothing Then
                Dim int_offset_rows As Long, int_offset_cols As Long
                int_offset_rows = butlSeries.name.Row - butlSeries.XValues.Cells(1, 1).Row
                int_offset_cols = butlSeries.name.Column - butlSeries.XValues.Cells(1, 1).Column

                Set butlSeries.name = butlSeries.Values.Cells(1, 1).Offset(int_offset_rows, int_offset_cols)
            End If

            butlSeries.UpdateSeriesWithNewValues

        Next chtSeries

        ''need to flip axis labels if they exist
        ''three cases: X only, Y only, X and Y

        If cht.Axes(xlCategory).HasTitle And Not cht.Axes(xlValue).HasTitle Then

            cht.Axes(xlValue).HasTitle = True
            cht.Axes(xlValue).AxisTitle.Text = cht.Axes(xlCategory).AxisTitle.Text
            cht.Axes(xlCategory).HasTitle = False

        ElseIf Not cht.Axes(xlCategory).HasTitle And cht.Axes(xlValue).HasTitle Then
            cht.Axes(xlCategory).HasTitle = True
            cht.Axes(xlCategory).AxisTitle.Text = cht.Axes(xlValue).AxisTitle.Text
            cht.Axes(xlValue).HasTitle = False
            
        ElseIf cht.Axes(xlCategory).HasTitle And cht.Axes(xlValue).HasTitle Then
            Dim swapText As String

            swapText = cht.Axes(xlCategory).AxisTitle.Text

            cht.Axes(xlCategory).AxisTitle.Text = cht.Axes(xlValue).AxisTitle.Text
            cht.Axes(xlValue).AxisTitle.Text = swapText

        End If

        Set butlSeriesies = Nothing

    Next chtObj

End Sub


Sub ChartMergeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartMergeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Merges all selected charts into a single chart
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim sel As Variant
    Dim firstChart As Chart

    Dim isFirstChart As Boolean
    isFirstChart = True
    
    Application.ScreenUpdating = False
    
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
    
        Set cht = chtObj.Chart
        If isFirstChart Then
            Set firstChart = cht
            isFirstChart = False
        Else
            Dim chtSeries As series
            For Each chtSeries In cht.SeriesCollection

                Dim chtNewSeries As series
                Dim butlSeries As New bUTLChartSeries

                butlSeries.UpdateFromChartSeries chtSeries
                Set chtNewSeries = butlSeries.AddSeriesToChart(firstChart)

                chtNewSeries.MarkerSize = chtSeries.MarkerSize
                chtNewSeries.MarkerStyle = chtSeries.MarkerStyle

                chtSeries.Delete

            Next chtSeries

            chtObj.Delete

        End If
    Next chtObj
    
    Application.ScreenUpdating = True

End Sub


Sub ChartSplitSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartSplitSeries
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Take all series from selected charts and puts them in their own charts
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    Dim cht As Chart

    Dim chtSeries As series
    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim chtObjNew As ChartObject
            Set chtObjNew = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim chtSeriesNew As series
            Dim butlSeries As New bUTLChartSeries

            butlSeries.UpdateFromChartSeries chtSeries
            Set chtSeriesNew = butlSeries.AddSeriesToChart(chtObjNew.Chart)

            chtSeriesNew.MarkerSize = chtSeries.MarkerSize
            chtSeriesNew.MarkerStyle = chtSeries.MarkerStyle

            chtSeries.Delete

        Next chtSeries


        chtObj.Delete

    Next chtObj
End Sub

