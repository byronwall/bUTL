Attribute VB_Name = "Chart_Processing"
Option Explicit

Public Sub Chart_CreateChartWithSeriesForEachColumn()
    'will create a chart that includes a series with no x value for each column

    Dim rng_data As Range
    Set rng_data = GetInputOrSelection("Select chart data")

    'create a chart
    Dim chtObj As ChartObject
    Set chtObj = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)
    
    chtObj.Chart.ChartType = xlXYScatter

    Dim rng_col As Range
    For Each rng_col In rng_data.Columns

        Dim rng_chart As Range
        Set rng_chart = RangeEnd(rng_col.Cells(1, 1), xlDown)
        
        Dim b_ser As New bUTLChartSeries
        Set b_ser.Values = rng_chart
        
        b_ser.AddSeriesToChart chtObj.Chart
    Next

End Sub

Public Sub Chart_CopyToSheet()

    Dim chtObj As ChartObject
    
    Dim obj_all As Object
    Set obj_all = Selection
    
    Dim msg_newSheet As VbMsgBoxResult
    msg_newSheet = MsgBox("New sheet?", vbYesNo, "New sheet?")
    
    Dim sht_out As Worksheet
    If msg_newSheet = vbYes Then
        Set sht_out = Worksheets.Add()
    Else
        Set sht_out = Application.InputBox("Pick a cell on a sheet", "Pick sheet", Type:=8).Parent
    End If
    
    For Each chtObj In Chart_GetObjectsFromObject(obj_all)
        chtObj.Copy
        
        sht_out.Paste
    Next
    
    sht_out.Activate
End Sub

Sub Chart_SortSeriesByName()
    'this will sort series by names
    Dim chtObj As ChartObject
    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        'uses a simple bubble sort but it works... shouldn't have 1000 series anyways
        Dim int_chart1 As Long
        Dim int_chart2 As Long
        For int_chart1 = 1 To chtObj.Chart.SeriesCollection.count
            For int_chart2 = (int_chart1 + 1) To chtObj.Chart.SeriesCollection.count

                Dim b_ser1 As New bUTLChartSeries
                Dim b_ser2 As New bUTLChartSeries

                b_ser1.UpdateFromChartSeries chtObj.Chart.SeriesCollection(int_chart1)
                b_ser2.UpdateFromChartSeries chtObj.Chart.SeriesCollection(int_chart2)

                If b_ser1.name.Value > b_ser2.name.Value Then
                    Dim int_num As Long
                    int_num = b_ser2.SeriesNumber
                    b_ser2.SeriesNumber = b_ser1.SeriesNumber
                    b_ser1.SeriesNumber = int_num

                    b_ser2.UpdateSeriesWithNewValues
                    b_ser1.UpdateSeriesWithNewValues
                End If
            Next
        Next
    Next
End Sub


Sub Chart_TimeSeries(rng_dates As Range, rng_data As Range, rng_titles As Range)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper Sub to create a set of charts with the same x axis and varying y
    '---------------------------------------------------------------------------------------
    '
    Application.ScreenUpdating = False

    Dim int_counter As Long
    int_counter = 1

    Dim rng_title As Range
    Dim rng_col As Range

    For Each rng_title In rng_titles

        Dim chtObj As ChartObject
        Set chtObj = ActiveSheet.ChartObjects.Add(int_counter * 300, 0, 300, 300)

        Dim cht As Chart
        Set cht = chtObj.Chart
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

    Application.ScreenUpdating = True
End Sub


Sub Chart_TimeSeries_FastCreation()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries_FastCreation
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : this will create a fast set of charts from a block of data
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
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



Sub CreateMultipleTimeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : CreateMultipleTimeSeries
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Entry point from Ribbon to create a set of time series charts
    '---------------------------------------------------------------------------------------
    '
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


Sub RemoveZeroValueDataLabel()
    '---------------------------------------------------------------------------------------
    ' Procedure : RemoveZeroValueDataLabel
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Code deletes data labels that have 0 value
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
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
        Dim i As Long
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

