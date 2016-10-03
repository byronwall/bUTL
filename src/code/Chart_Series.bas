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

        Dim ser As series

        Dim i As Long
        i = 1

        For Each ser In chtObj.Chart.SeriesCollection

            Dim b_ser As New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            'clear out old ones
            Dim j As Long
            For j = 1 To ser.Trendlines.count
                ser.Trendlines(j).Delete
            Next j

            ser.MarkerBackgroundColor = Chart_GetColor(i)

            Dim trend As Trendline
            Set trend = ser.Trendlines.Add()
            trend.Type = xlLinear
            trend.Border.Color = ser.MarkerBackgroundColor
            
            '2015 11 06 test to avoid error without name
            '2015 12 07 dealing with multi-cell Names
            'TODO: handle if the name is not a range also
            If Not b_ser.name Is Nothing Then
                trend.name = b_ser.name.Cells(1, 1).Value
            End If

            trend.DisplayEquation = True
            trend.DisplayRSquared = True
            trend.DataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Chart_GetColor(i)

            i = i + 1
        Next ser

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

        Dim ser As series

        'get each series
        For Each ser In chtObj.Chart.SeriesCollection

            'create the bUTL obj and manipulate series ranges
            Dim b_ser As New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            If Not b_ser.XValues Is Nothing Then
                ser.XValues = RangeEnd(b_ser.XValues.Cells(1), xlDown)
            End If
            ser.Values = RangeEnd(b_ser.Values.Cells(1), xlDown)

        Next ser

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


Sub Chart_GoToYRange()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GoToYRange
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Selects the y values used for the series
    '---------------------------------------------------------------------------------------
    '
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

        Dim ser As series
        For Each ser In chtObj.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In ser.Trendlines
                trend.Delete
            Next trend

        Next ser

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

        Dim series As series

        For Each series In chtObj.Chart.SeriesCollection

            Dim trend As Trendline

            For Each trend In series.Trendlines
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
                Dim int_offset_rows As Long, int_offset_cols As Long
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
    Dim cht_first As Chart

    Dim bool_first As Boolean
    bool_first = True
    
    Application.ScreenUpdating = False
    
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
    
        Set cht = chtObj.Chart
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
    Dim sel As Variant

    Dim ser As series
    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        For Each ser In chtObj.Chart.SeriesCollection

            Dim chtObj_new As ChartObject
            Set chtObj_new = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)

            Dim ser_new As series
            Dim b_ser As New bUTLChartSeries

            b_ser.UpdateFromChartSeries ser
            Set ser_new = b_ser.AddSeriesToChart(chtObj_new.Chart)

            ser_new.MarkerSize = ser.MarkerSize
            ser_new.MarkerStyle = ser.MarkerStyle

            ser.Delete

        Next ser


        chtObj.Delete

    Next chtObj
End Sub

