Attribute VB_Name = "Chart_Processing"
Option Explicit

Public Sub Chart_CreateChartWithSeriesForEachColumn()
    'will create a chart that includes a series with no x value for each column

    Dim dataRange As Range
    Set dataRange = GetInputOrSelection("Select chart data")

    'create a chart
    Dim targetObject As ChartObject
    Set targetObject = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)
    
    targetObject.Chart.ChartType = xlXYScatter

    Dim targetColumn As Range
    For Each targetColumn In dataRange.Columns

        Dim chartDataRange As Range
        Set chartDataRange = RangeEnd(targetColumn.Cells(1, 1), xlDown)
        
        Dim butlSeries As New bUTLChartSeries
        Set butlSeries.Values = chartDataRange
        
        butlSeries.AddSeriesToChart targetObject.Chart
    Next targetColumn

End Sub

Public Sub Chart_CopyToSheet()

    Dim targetObject As ChartObject
    
    Dim selectedObject As Object
    Set selectedObject = Selection
    
    Dim newSheetResult As VbMsgBoxResult
    newSheetResult = MsgBox("Create a new sheet?", vbYesNo, "New sheet?")
    
    Dim targetSheet As Worksheet
    If newSheetResult = vbYes Then
        Set targetSheet = Worksheets.Add()
    Else: Set targetSheet = Application.InputBox("Pick a cell on an existing sheet", "Pick sheet", Type:=8).Parent
    End If
    
    For Each targetObject In Chart_GetObjectsFromObject(selectedObject)
        targetObject.Copy
        targetSheet.Paste
    Next targetObject
    
    targetSheet.Activate
End Sub

Public Sub Chart_SortSeriesByName()
    'this will sort series by names
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        'uses a simple bubble sort but it works... shouldn't have 1000 series anyways
        Dim firstChartIndex As Long
        Dim secondChartIndex As Long
        For firstChartIndex = 1 To targetObject.Chart.SeriesCollection.count
            For secondChartIndex = (firstChartIndex + 1) To targetObject.Chart.SeriesCollection.count

                Dim butlSeries1 As New bUTLChartSeries
                Dim butlSeries2 As New bUTLChartSeries

                butlSeries1.UpdateFromChartSeries targetObject.Chart.SeriesCollection(firstChartIndex)
                butlSeries2.UpdateFromChartSeries targetObject.Chart.SeriesCollection(secondChartIndex)

                If butlSeries1.name.Value > butlSeries2.name.Value Then
                    Dim indexSeriesSwap As Long
                    indexSeriesSwap = butlSeries2.SeriesNumber
                    butlSeries2.SeriesNumber = butlSeries1.SeriesNumber
                    butlSeries1.SeriesNumber = indexSeriesSwap
                    butlSeries2.UpdateSeriesWithNewValues
                    butlSeries1.UpdateSeriesWithNewValues
                End If
                
            Next secondChartIndex
        Next firstChartIndex
    Next targetObject
End Sub


Public Sub Chart_TimeSeries(ByVal rangeOfDates As Range, ByVal dataRange As Range, ByVal rangeOfTitles As Range)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper Sub to create a set of charts with the same x axis and varying y
    '---------------------------------------------------------------------------------------
    '
    Application.ScreenUpdating = False
    Const MARKER_SIZE As Long = 3
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(200, 200, 200)
    Dim chartIndex As Long
    chartIndex = 1

    Dim titleRange As Range
    Dim targetColumn As Range

    For Each titleRange In rangeOfTitles

        Dim targetObject As ChartObject
        Set targetObject = ActiveSheet.ChartObjects.Add(chartIndex * 300, 0, 300, 300)

        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        targetChart.ChartType = xlXYScatterLines
        targetChart.HasTitle = True
        targetChart.Legend.Delete

        Dim targetAxis As Axis
        Set targetAxis = targetChart.Axes(xlValue)
        targetAxis.MajorGridlines.Border.Color = majorGridlineColor

        Dim targetSeries As series
        Dim butlSeries As New bUTLChartSeries

        Set butlSeries.XValues = rangeOfDates
        Set butlSeries.Values = dataRange.Columns(chartIndex)
        Set butlSeries.name = titleRange

        Set targetSeries = butlSeries.AddSeriesToChart(targetChart)

        targetSeries.MarkerSize = MARKER_SIZE
        targetSeries.MarkerStyle = xlMarkerStyleCircle

        chartIndex = chartIndex + 1

    Next titleRange

    Application.ScreenUpdating = True
End Sub


Public Sub Chart_TimeSeries_FastCreation()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries_FastCreation
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : this will create a fast set of charts from a block of data
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim rangeOfDates As Range
    Dim dataRange As Range
    Dim rangeOfTitles As Range

    'dates are in B4 and down
    Set rangeOfDates = RangeEnd_Boundary(Range("B4"), xlDown)

    'data starts in C4, down and over
    Set dataRange = RangeEnd_Boundary(Range("C4"), xlDown, xlToRight)

    'titels are C2 and over
    Set rangeOfTitles = RangeEnd_Boundary(Range("C2"), xlToRight)

    Chart_TimeSeries rangeOfDates, dataRange, rangeOfTitles
    ChartDefaultFormat
    Chart_GridOfCharts

End Sub



Public Sub CreateMultipleTimeSeries()
    '---------------------------------------------------------------------------------------
    ' Procedure : CreateMultipleTimeSeries
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Entry point from Ribbon to create a set of time series charts
    '---------------------------------------------------------------------------------------
    '
    Dim rangeOfDates As Range
    Dim dataRange As Range
    Dim rangeOfTitles As Range

    On Error GoTo CreateMultipleTimeSeries_Error

    DeleteAllCharts

    Set rangeOfDates = Application.InputBox("Select date range", Type:=8)
    Set dataRange = Application.InputBox("Select data", Type:=8)
    Set rangeOfTitles = Application.InputBox("Select titles", Type:=8)

    Chart_TimeSeries rangeOfDates, dataRange, rangeOfTitles

    On Error GoTo 0
    Exit Sub

CreateMultipleTimeSeries_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & "), likely due to Range selection."

End Sub


Public Sub RemoveZeroValueDataLabel()
    '---------------------------------------------------------------------------------------
    ' Procedure : RemoveZeroValueDataLabel
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Code deletes data labels that have 0 value
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    'uses the ActiveChart, be sure a chart is selected
    Dim targetChart As Chart
    Set targetChart = ActiveChart

    Dim targetSeries As series
    For Each targetSeries In targetChart.SeriesCollection

        Dim seriesValues As Variant
        seriesValues = targetSeries.Values

        'include this line if you want to reestablish labels before deleting
        targetSeries.ApplyDataLabels xlDataLabelsShowLabel, , , , True, False, False, False, False

        'loop through values and delete 0-value labels
        Dim pointIndex As Long
        For pointIndex = LBound(seriesValues) To UBound(seriesValues)
            If seriesValues(pointIndex) = 0 Then
                With targetSeries.Points(pointIndex)
                    If .HasDataLabel Then .DataLabel.Delete
                End With
            End If
        Next pointIndex
    Next targetSeries
End Sub

