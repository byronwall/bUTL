Attribute VB_Name = "Chart_Series"
Option Explicit

Public Sub Chart_AddTrendlineToSeriesAndColor()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AddTrendlineToSeriesAndColor
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds a trendline to each series in all charts
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim chartIndex As Long
        chartIndex = 1
        
        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            'clear out old ones
            Dim j As Long
            For j = 1 To targetSeries.Trendlines.count
                targetSeries.Trendlines(j).Delete
            Next j

            targetSeries.MarkerBackgroundColor = Chart_GetColor(chartIndex)

            Dim newTrendline As Trendline
            Set newTrendline = targetSeries.Trendlines.Add()
            newTrendline.Type = xlLinear
            newTrendline.Border.Color = targetSeries.MarkerBackgroundColor
            
            '2015 11 06 test to avoid error without name
            '2015 12 07 dealing with multi-cell Names
            'TODO: handle if the name is not a range also
            If Not butlSeries.name Is Nothing Then
                newTrendline.name = butlSeries.name.Cells(1, 1).Value
            End If

            newTrendline.DisplayEquation = True
            newTrendline.DisplayRSquared = True
            newTrendline.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(chartIndex)

            chartIndex = chartIndex + 1
        Next targetSeries

    Next targetObject
End Sub


Public Sub Chart_ExtendSeriesToRanges()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ExtendSeriesToRanges
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Extends the underlying data for a series to go to the end of its current Range
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        'get each series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            If Not butlSeries.XValues Is Nothing Then
                targetSeries.XValues = RangeEnd(butlSeries.XValues.Cells(1), xlDown)
            End If
            targetSeries.Values = RangeEnd(butlSeries.Values.Cells(1), xlDown)

        Next targetSeries
    Next targetObject
End Sub


Public Sub Chart_GoToXRange()
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


Public Sub Chart_GoToYRange()
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


Public Sub Chart_RemoveTrendlines()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_RemoveTrendlines
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Remove all trendlines from a chart
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newTrendline As Trendline
            For Each newTrendline In targetSeries.Trendlines
                newTrendline.Delete
            Next newTrendline
        Next targetSeries
    Next targetObject
End Sub


Public Sub Chart_RerangeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_RerangeSeries
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Entry point for an interface to help rerange series
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetForm As New form_chtSeries
    targetForm.Show

End Sub


Public Sub Chart_TrendlinesToAverage()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TrendlinesToAverage
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a trendline using a moving average instead of linear
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newTrendline As Trendline

            For Each newTrendline In targetSeries.Trendlines
                newTrendline.Type = xlMovingAvg
                newTrendline.Period = 15
                newTrendline.Format.Line.Weight = 2
            Next
        Next
    Next

End Sub


Public Sub ChartFlipXYValues()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartFlipXYValues
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Flips the x/y ranges for each series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Set targetChart = targetObject.Chart

        Dim butlSeriesies As New Collection
        Dim butlSeries As bUTLChartSeries
        
        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            Dim dummyRange As Range

            Set dummyRange = butlSeries.Values
            Set butlSeries.Values = butlSeries.XValues
            Set butlSeries.XValues = dummyRange

            'need to change the series name also
            'assume that title is same offset
            'code blocked for now
            If False And Not butlSeries.name Is Nothing Then
                Dim rowsOffset As Long, columnsOffset As Long
                rowsOffset = butlSeries.name.Row - butlSeries.XValues.Cells(1, 1).Row
                columnsOffset = butlSeries.name.Column - butlSeries.XValues.Cells(1, 1).Column

                Set butlSeries.name = butlSeries.Values.Cells(1, 1).Offset(rowsOffset, columnsOffset)
            End If

            butlSeries.UpdateSeriesWithNewValues

        Next targetSeries

        ''need to flip axis labels if they exist
        ''three cases: X only, Y only, X and Y

        If targetChart.Axes(xlCategory).HasTitle And Not targetChart.Axes(xlValue).HasTitle Then

            targetChart.Axes(xlValue).HasTitle = True
            targetChart.Axes(xlValue).AxisTitle.Text = targetChart.Axes(xlCategory).AxisTitle.Text
            targetChart.Axes(xlCategory).HasTitle = False

        ElseIf Not targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then
            targetChart.Axes(xlCategory).HasTitle = True
            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text
            targetChart.Axes(xlValue).HasTitle = False
            
        ElseIf targetChart.Axes(xlCategory).HasTitle And targetChart.Axes(xlValue).HasTitle Then
            Dim swapText As String

            swapText = targetChart.Axes(xlCategory).AxisTitle.Text

            targetChart.Axes(xlCategory).AxisTitle.Text = targetChart.Axes(xlValue).AxisTitle.Text
            targetChart.Axes(xlValue).AxisTitle.Text = swapText

        End If

        Set butlSeriesies = Nothing

    Next targetObject

End Sub


Public Sub ChartMergeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartMergeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Merges all selected charts into a single chart
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart
    Dim firstChart As Chart

    Dim isFirstChart As Boolean
    isFirstChart = True
    
    Application.ScreenUpdating = False
    
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
    
        Set targetChart = targetObject.Chart
        If isFirstChart Then
            Set firstChart = targetChart
            isFirstChart = False
        Else
            Dim targetSeries As series
            For Each targetSeries In targetChart.SeriesCollection

                Dim newChartSeries As series
                Dim butlSeries As New bUTLChartSeries

                butlSeries.UpdateFromChartSeries targetSeries
                Set newChartSeries = butlSeries.AddSeriesToChart(firstChart)

                newChartSeries.MarkerSize = targetSeries.MarkerSize
                newChartSeries.MarkerStyle = targetSeries.MarkerStyle

                targetSeries.Delete

            Next targetSeries

            targetObject.Delete

        End If
    Next targetObject
    
    Application.ScreenUpdating = True

End Sub


Public Sub ChartSplitSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartSplitSeries
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Take all series from selected charts and puts them in their own charts
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart

    Dim targetSeries As series
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim newChartObject As ChartObject
            Set newChartObject = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim newChartSeries As series
            Dim butlSeries As New bUTLChartSeries

            butlSeries.UpdateFromChartSeries targetSeries
            Set newChartSeries = butlSeries.AddSeriesToChart(newChartObject.Chart)

            newChartSeries.MarkerSize = targetSeries.MarkerSize
            newChartSeries.MarkerStyle = targetSeries.MarkerStyle

            targetSeries.Delete

        Next targetSeries


        targetObject.Delete

    Next targetObject
End Sub

