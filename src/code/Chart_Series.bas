Attribute VB_Name = "Chart_Series"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : Chart_Series
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains charting code related to managing series
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : Chart_AddTrendlineToSeriesAndColor
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds a trendline to each series in all charts
'---------------------------------------------------------------------------------------
'
Sub Chart_AddTrendlineToSeriesAndColor()

    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series

        Dim i As Long
        i = 1

        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim ButlSeries As New bUTLChartSeries
            ButlSeries.UpdateFromChartSeries mySeries

            'clear out old ones
            Dim j As Long
            For j = 1 To mySeries.Trendlines.count
                mySeries.Trendlines(j).Delete
            Next j

            mySeries.MarkerBackgroundColor = Chart_GetColor(i)

            Dim trend As Trendline
            Set trend = mySeries.Trendlines.Add()
            trend.Type = xlLinear
            trend.Border.Color = mySeries.MarkerBackgroundColor
            
            '2015 11 06 test to avoid error without name
            If Not ButlSeries.name Is Nothing Then
                trend.name = ButlSeries.name
            End If

            trend.DisplayEquation = True
            trend.DisplayRSquared = True
            trend.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(i)

            i = i + 1
        Next mySeries

    Next myChartObject
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_ExtendSeriesToRanges
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Extends the underlying data for a series to go to the end of its current Range
'---------------------------------------------------------------------------------------
'
Sub Chart_ExtendSeriesToRanges()

    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series

        'get each series
        For Each mySeries In myChartObject.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim ButlSeries As New bUTLChartSeries
            ButlSeries.UpdateFromChartSeries mySeries

            If Not ButlSeries.XValues Is Nothing Then
                mySeries.XValues = RangeEnd(ButlSeries.XValues.Cells(1), xlDown)
            End If
            mySeries.Values = RangeEnd(ButlSeries.Values.Cells(1), xlDown)

        Next mySeries

    Next myChartObject


End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_GoToXRange
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Selects the x value range that is used for the series
'---------------------------------------------------------------------------------------
'
Sub Chart_GoToXRange()



    If TypeName(Selection) = "Series" Then
        Dim ButlSeries As New bUTLChartSeries
        ButlSeries.UpdateFromChartSeries Selection

        ButlSeries.XValues.Parent.Activate
        ButlSeries.XValues.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_GoToYRange
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Selects the y values used for the series
'---------------------------------------------------------------------------------------
'
Sub Chart_GoToYRange()



    If TypeName(Selection) = "Series" Then
        Dim ButlSeries As New bUTLChartSeries
        ButlSeries.UpdateFromChartSeries Selection

        ButlSeries.Values.Parent.Activate
        ButlSeries.Values.Activate
    Else
        MsgBox "Select a series in order to use this."
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_RemoveTrendlines
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Remove all trendlines from a chart
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub Chart_RemoveTrendlines()

    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series
        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In mySeries.Trendlines
                trend.Delete
            Next trend

        Next mySeries

    Next myChartObject
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_RerangeSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Entry point for an interface to help rerange series
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub Chart_RerangeSeries()

    Dim frm As New form_chtSeries
    frm.Show

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_TrendlinesToAverage
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates a trendline using a moving average instead of linear
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub Chart_TrendlinesToAverage()
    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series

        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In mySeries.Trendlines
                trend.Type = xlMovingAvg
                trend.Period = 15
                trend.Format.Line.Weight = 2
            Next
        Next
    Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartFlipXYValues
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Flips the x/y ranges for each series
'---------------------------------------------------------------------------------------
'
Sub ChartFlipXYValues()

    Dim myChartObject As ChartObject
    Dim myChart As Chart
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
        Set myChart = myChartObject.Chart

        Dim mySeries As series

        Dim ButlSeriesies As New Collection
        Dim ButlSeries As bUTLChartSeries

        For Each mySeries In myChart.SeriesCollection
            Set ButlSeries = New bUTLChartSeries
            ButlSeries.UpdateFromChartSeries mySeries

            Dim rng_dummy As Range

            Set rng_dummy = ButlSeries.Values
            Set ButlSeries.Values = ButlSeries.XValues
            Set ButlSeries.XValues = rng_dummy

            'need to change the series name also
            'assume that title is same offset
            'code blocked for now
            If False And Not ButlSeries.name Is Nothing Then
                Dim int_offset_rows As Long, int_offset_cols As Long
                int_offset_rows = ButlSeries.name.row - ButlSeries.XValues.Cells(1, 1).row
                int_offset_cols = ButlSeries.name.Column - ButlSeries.XValues.Cells(1, 1).Column

                Set ButlSeries.name = ButlSeries.Values.Cells(1, 1).Offset(int_offset_rows, int_offset_cols)
            End If

            ButlSeries.UpdateSeriesWithNewValues

        Next mySeries

        ''need to flip axis labels if they exist


        ''three cases: X only, Y only, X and Y

        If myChart.Axes(xlCategory).HasTitle And Not myChart.Axes(xlValue).HasTitle Then

            myChart.Axes(xlValue).HasTitle = True
            myChart.Axes(xlValue).AxisTitle.Text = myChart.Axes(xlCategory).AxisTitle.Text
            myChart.Axes(xlCategory).HasTitle = False

        ElseIf Not myChart.Axes(xlCategory).HasTitle And myChart.Axes(xlValue).HasTitle Then
            myChart.Axes(xlCategory).HasTitle = True
            myChart.Axes(xlCategory).AxisTitle.Text = myChart.Axes(xlValue).AxisTitle.Text
            myChart.Axes(xlValue).HasTitle = False
        ElseIf myChart.Axes(xlCategory).HasTitle And myChart.Axes(xlValue).HasTitle Then
            Dim tempString As String

            tempString = myChart.Axes(xlCategory).AxisTitle.Text

            myChart.Axes(xlCategory).AxisTitle.Text = myChart.Axes(xlValue).AxisTitle.Text
            myChart.Axes(xlValue).AxisTitle.Text = tempString

        End If

        Set ButlSeriesies = Nothing

    Next myChartObject

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartMergeSeries
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Merges all selected charts into a single chart
'---------------------------------------------------------------------------------------
'
Sub ChartMergeSeries()

    Dim myChartObject As ChartObject
    Dim myChart As Chart
  
    Dim firstChart As Chart

    Dim first As Boolean
    first = True
    
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
    
        Set myChart = myChartObject.Chart
        If first Then
            Set firstChart = myChart
            first = False
        Else
            Dim mySeries As series
            For Each mySeries In myChart.SeriesCollection

                Dim newSeries As series
                Dim ButlSeries As New bUTLChartSeries

                ButlSeries.UpdateFromChartSeries mySeries
                Set newSeries = ButlSeries.AddSeriesToChart(firstChart)

                newSeries.MarkerSize = mySeries.MarkerSize
                newSeries.MarkerStyle = mySeries.MarkerStyle

                mySeries.Delete

            Next mySeries

            myChartObject.Delete

        End If
    Next myChartObject

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartSplitSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Take all series from selected charts and puts them in their own charts
'---------------------------------------------------------------------------------------
'
Sub ChartSplitSeries()

    Dim myChartObject As ChartObject

  

    Dim mySeries As series
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim newChartObject As ChartObject
            Set newChartObject = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim newSeries As series
            Dim ButlSeries As New bUTLChartSeries

            ButlSeries.UpdateFromChartSeries mySeries
            Set newSeries = ButlSeries.AddSeriesToChart(newChartObject.Chart)

            newSeries.MarkerSize = mySeries.MarkerSize
            newSeries.MarkerStyle = mySeries.MarkerStyle

            mySeries.Delete

        Next mySeries


        myChartObject.Delete

    Next myChartObject
End Sub

