Attribute VB_Name = "Chart_Processing"
'---------------------------------------------------------------------------------------
' Module    : Chart_Processing
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains some of the heavy lifting processing code for charts
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Chart_TimeSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper Sub to create a set of charts with the same x axis and varying y
'---------------------------------------------------------------------------------------
'
Sub Chart_TimeSeries(rng_dates As Range, rng_data As Range, rng_titles As Range)

    Dim int_counter As Integer
    int_counter = 1

    Dim rng_title As Range
    Dim rng_col As Range

    For Each rng_title In rng_titles

        Dim cht_obj As ChartObject
        Set cht_obj = ActiveSheet.ChartObjects.Add(int_counter * 300, 0, 300, 300)

        Dim cht As Chart
        Set cht = cht_obj.Chart
        cht.ChartType = xlXYScatterLines
        cht.HasTitle = True
        cht.Legend.Delete

        Dim ax As Axis
        Set ax = cht.Axes(xlValue)
        ax.MajorGridlines.Border.Color = RGB(200, 200, 200)

        Dim ser As series
        Dim b_ser As New bUTLChartSeries

        Set b_ser.XValues = rng_dates
        Set b_ser.Values = rng_data.Columns(int_counter)
        Set b_ser.name = rng_title

        Set ser = b_ser.AddSeriesToChart(cht)

        ser.MarkerSize = 3
        ser.MarkerStyle = xlMarkerStyleCircle

        int_counter = int_counter + 1

    Next rng_title
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_TimeSeries_FastCreation
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : this will create a fast set of charts from a block of data
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub Chart_TimeSeries_FastCreation()

    Dim rng_dates As Range
    Dim rng_data As Range
    Dim rng_titles As Range

    'dates are in B4 and down
    Set rng_dates = RangeEnd_Boundary(Range("B4"), xlDown)

    'data starts in C4, down and over
    Set rng_data = RangeEnd_Boundary(Range("C4"), xlDown, xlToRight)

    'titels are C2 and over
    Set rng_titles = RangeEnd_Boundary(Range("C2"), xlToRight)

    Chart_TimeSeries rng_dates, rng_data, rng_titles
    ChartDefaultFormat
    Chart_GridOfCharts

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateMultipleTimeSeries
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Entry point from Ribbon to create a set of time series charts
'---------------------------------------------------------------------------------------
'
Sub CreateMultipleTimeSeries()

    Dim rng_dates As Range
    Dim rng_data As Range
    Dim rng_titles As Range

    On Error GoTo CreateMultipleTimeSeries_Error

    DeleteAllCharts

    Set rng_dates = Application.InputBox("Select date range", Type:=8)
    Set rng_data = Application.InputBox("Select data", Type:=8)
    Set rng_titles = Application.InputBox("Select titles", Type:=8)

    Chart_TimeSeries rng_dates, rng_data, rng_titles

    On Error GoTo 0
    Exit Sub

CreateMultipleTimeSeries_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & "), likely due to Range selection."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : RemoveZeroValueDataLabel
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Code deletes data labels that have 0 value
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub RemoveZeroValueDataLabel()

'uses the ActiveChart, be sure a chart is selected
    Dim cht As Chart
    Set cht = ActiveChart

    Dim ser As series
    For Each ser In cht.SeriesCollection

        Dim vals As Variant
        vals = ser.Values

        'include this line if you want to reestablish labels before deleting
        ser.ApplyDataLabels xlDataLabelsShowLabel, , , , True, False, False, False, False

        'loop through values and delete 0-value labels
        Dim i As Integer
        For i = LBound(vals) To UBound(vals)
            If vals(i) = 0 Then
                With ser.Points(i)
                    If .HasDataLabel Then
                        .DataLabel.Delete
                    End If
                End With
            End If
        Next i
    Next ser
End Sub

