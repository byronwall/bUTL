Attribute VB_Name = "Chart_Series"
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

    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series

        Dim i As Integer
        i = 1

        For Each ser In cht_obj.Chart.SeriesCollection

            Dim b_ser As New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            'clear out old ones
            Dim j As Integer
            For j = 1 To ser.Trendlines.count
                ser.Trendlines(j).Delete
            Next j

            ser.MarkerBackgroundColor = Chart_GetColor(i)

            Dim trend As Trendline
            Set trend = ser.Trendlines.Add()
            trend.Type = xlLinear
            trend.Border.Color = ser.MarkerBackgroundColor
            trend.name = b_ser.name

            trend.DisplayEquation = True
            trend.DisplayRSquared = True
            trend.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(i)

            i = i + 1
        Next ser

    Next cht_obj
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_ExtendSeriesToRanges
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Extends the underlying data for a series to go to the end of its current Range
'---------------------------------------------------------------------------------------
'
Sub Chart_ExtendSeriesToRanges()

    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series

        'get each series
        For Each ser In cht_obj.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim b_ser As New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            ser.XValues = RangeEnd(b_ser.XValues, xlDown)
            ser.Values = RangeEnd(b_ser.Values, xlDown)

        Next ser

    Next cht_obj


End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_GoToXRange
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Selects the x value range that is used for the series
'---------------------------------------------------------------------------------------
'
Sub Chart_GoToXRange()

    Dim ser As series

    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.XValues.Parent.Activate
        b.XValues.Activate
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

    Dim ser As series

    If TypeName(Selection) = "Series" Then
        Dim b As New bUTLChartSeries
        b.UpdateFromChartSeries Selection

        b.Values.Parent.Activate
        b.Values.Activate
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

    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series
        For Each ser In cht_obj.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In ser.Trendlines
                trend.Delete
            Next trend

        Next ser

    Next cht_obj
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
    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        Dim series As series

        For Each series In cht_obj.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In series.Trendlines
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

    Dim cht_obj As ChartObject
    Dim cht As Chart
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Set cht = cht_obj.Chart

        Dim ser As series

        Dim b_series As New Collection
        Dim b_ser As bUTLChartSeries

        For Each ser In cht.SeriesCollection
            Set b_ser = New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            Dim rng_dummy As Range

            Set rng_dummy = b_ser.Values
            Set b_ser.Values = b_ser.XValues
            Set b_ser.XValues = rng_dummy

            'need to change the series name also
            'assume that title is same offset
            'code blocked for now
            If False And Not b_ser.name Is Nothing Then
                Dim int_offset_rows As Integer, int_offset_cols As Integer
                int_offset_rows = b_ser.name.Row - b_ser.XValues.Cells(1, 1).Row
                int_offset_cols = b_ser.name.Column - b_ser.XValues.Cells(1, 1).Column

                Set b_ser.name = b_ser.Values.Cells(1, 1).Offset(int_offset_rows, int_offset_cols)
            End If

            b_ser.UpdateSeriesWithNewValues

        Next ser

        ''need to flip axis labels if they exist
        Dim dummy_title As AxisTitle

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
            Dim dummy_text As String

            dummy_text = cht.Axes(xlCategory).AxisTitle.Text

            cht.Axes(xlCategory).AxisTitle.Text = cht.Axes(xlValue).AxisTitle.Text
            cht.Axes(xlValue).AxisTitle.Text = dummy_text

        End If

        Set b_series = Nothing

    Next cht_obj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartMergeSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Merges all selected charts into a single chart
'---------------------------------------------------------------------------------------
'
Sub ChartMergeSeries()

    Dim cht_obj As ChartObject
    Dim cht As Chart
    Dim sel As Variant
    Dim cht_first As Chart

    Dim bool_first As Boolean
    bool_first = True
    For Each cht_obj In Selection
        Set cht = cht_obj.Chart
        If bool_first Then
            Set cht_first = cht
            bool_first = False
        Else
            Dim ser As series
            For Each ser In cht.SeriesCollection

                Dim ser_new As series
                Dim b_ser As New bUTLChartSeries

                b_ser.UpdateFromChartSeries ser
                Set ser_new = b_ser.AddSeriesToChart(cht_first)

                ser_new.MarkerSize = ser.MarkerSize
                ser_new.MarkerStyle = ser.MarkerStyle

                ser.Delete

            Next ser

            cht_obj.Delete

        End If

    Next cht_obj



End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartSplitSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Take all series from selected charts and puts them in their own charts
'---------------------------------------------------------------------------------------
'
Sub ChartSplitSeries()

    Dim cht_obj As ChartObject
    Dim cht As Chart
    Dim sel As Variant

    Dim ser As series
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        For Each ser In cht_obj.Chart.SeriesCollection

            Dim cht_obj_new As ChartObject
            Set cht_obj_new = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim ser_new As series
            Dim b_ser As New bUTLChartSeries

            b_ser.UpdateFromChartSeries ser
            Set ser_new = b_ser.AddSeriesToChart(cht_obj_new.Chart)

            ser_new.MarkerSize = ser.MarkerSize
            ser_new.MarkerStyle = ser.MarkerStyle

            ser.Delete

        Next ser


        cht_obj.Delete

    Next cht_obj
End Sub

