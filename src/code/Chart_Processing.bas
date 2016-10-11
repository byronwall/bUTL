Attribute VB_Name = "Chart_Processing"
Option Explicit

Public Sub Chart_CreateChartWithSeriesForEachColumn()
    'will create a chart that includes a series with no x value for each column

    Dim rngData As Range
    Set rngData = GetInputOrSelection("Select chart data")

    'create a chart
    Dim chtObj As ChartObject
    Set chtObj = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)
    
    chtObj.Chart.ChartType = xlXYScatter

    Dim rngColumn As Range
    For Each rngColumn In rngData.Columns

        Dim rngChartData As Range
        Set rngChartData = RangeEnd(rngColumn.Cells(1, 1), xlDown)
        
        Dim butlSeries As New bUTLChartSeries
        Set butlSeries.Values = rngChartData
        
        butlSeries.AddSeriesToChart chtObj.Chart
    Next

End Sub

Public Sub Chart_CopyToSheet()

    Dim chtObj As ChartObject
    
    Dim objSelection As Object
    Set objSelection = Selection
    
    Dim newSheetResult As VbMsgBoxResult
    newSheetResult = MsgBox("New sheet?", vbYesNo, "New sheet?")
    
    Dim shtOutput As Worksheet
    If newSheetResult = vbYes Then
        Set shtOutput = Worksheets.Add()
    Else
        Set shtOutput = Application.InputBox("Pick a cell on a sheet", "Pick sheet", Type:=8).Parent
    End If
    
    For Each chtObj In Chart_GetObjectsFromObject(objSelection)
        chtObj.Copy
        
        shtOutput.Paste
    Next
    
    shtOutput.Activate
End Sub

Sub Chart_SortSeriesByName()
    'this will sort series by names
    Dim chtObj As ChartObject
    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        'uses a simple bubble sort but it works... shouldn't have 1000 series anyways
        Dim chtIndex1 As Long
        Dim chtIndex2 As Long
        For chtIndex1 = 1 To chtObj.Chart.SeriesCollection.count
            For chtIndex2 = (chtIndex1 + 1) To chtObj.Chart.SeriesCollection.count

                Dim butlSeries1 As New bUTLChartSeries
                Dim butlSeries2 As New bUTLChartSeries

                butlSeries1.UpdateFromChartSeries chtObj.Chart.SeriesCollection(chtIndex1)
                butlSeries2.UpdateFromChartSeries chtObj.Chart.SeriesCollection(chtIndex2)

                If butlSeries1.name.Value > butlSeries2.name.Value Then
                    Dim indexSeriesSwap As Long
                    indexSeriesSwap = butlSeries2.SeriesNumber
                    butlSeries2.SeriesNumber = butlSeries1.SeriesNumber
                    butlSeries1.SeriesNumber = indexSeriesSwap

                    butlSeries2.UpdateSeriesWithNewValues
                    butlSeries1.UpdateSeriesWithNewValues
                End If
            Next
        Next
    Next
End Sub


Sub Chart_TimeSeries(rngDates As Range, rngData As Range, rngTitles As Range)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_TimeSeries
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper Sub to create a set of charts with the same x axis and varying y
    '---------------------------------------------------------------------------------------
    '
    Application.ScreenUpdating = False

    Dim chartIndex As Long
    chartIndex = 1

    Dim rngTitle As Range
    Dim rngColumn As Range

    For Each rngTitle In rngTitles

        Dim chtObj As ChartObject
        Set chtObj = ActiveSheet.ChartObjects.Add(chartIndex * 300, 0, 300, 300)

        Dim cht As Chart
        Set cht = chtObj.Chart
        cht.ChartType = xlXYScatterLines
        cht.HasTitle = True
        cht.Legend.Delete

        Dim chtAxis As Axis
        Set chtAxis = cht.Axes(xlValue)
        chtAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)

        Dim chtSeries As series
        Dim butlSeries As New bUTLChartSeries

        Set butlSeries.XValues = rngDates
        Set butlSeries.Values = rngData.Columns(chartIndex)
        Set butlSeries.name = rngTitle

        Set chtSeries = butlSeries.AddSeriesToChart(cht)

        chtSeries.MarkerSize = 3
        chtSeries.MarkerStyle = xlMarkerStyleCircle

        chartIndex = chartIndex + 1

    Next rngTitle

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
    Dim rngDates As Range
    Dim rngData As Range
    Dim rngTitles As Range

    'dates are in B4 and down
    Set rngDates = RangeEnd_Boundary(Range("B4"), xlDown)

    'data starts in C4, down and over
    Set rngData = RangeEnd_Boundary(Range("C4"), xlDown, xlToRight)

    'titels are C2 and over
    Set rngTitles = RangeEnd_Boundary(Range("C2"), xlToRight)

    Chart_TimeSeries rngDates, rngData, rngTitles
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
    Dim rngDates As Range
    Dim rngData As Range
    Dim rngTitles As Range

    On Error GoTo CreateMultipleTimeSeries_Error

    DeleteAllCharts

    Set rngDates = Application.InputBox("Select date range", Type:=8)
    Set rngData = Application.InputBox("Select data", Type:=8)
    Set rngTitles = Application.InputBox("Select titles", Type:=8)

    Chart_TimeSeries rngDates, rngData, rngTitles

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

    Dim chtSeries As series
    For Each chtSeries In cht.SeriesCollection

        Dim chtValues As Variant
        chtValues = chtSeries.Values

        'include this line if you want to reestablish labels before deleting
        chtSeries.ApplyDataLabels xlDataLabelsShowLabel, , , , True, False, False, False, False

        'loop through values and delete 0-value labels
        Dim pointIndex As Long
        For pointIndex = LBound(chtValues) To UBound(chtValues)
            If chtValues(pointIndex) = 0 Then
                With chtSeries.Points(pointIndex)
                    If .HasDataLabel Then
                        .DataLabel.Delete
                    End If
                End With
            End If
        Next pointIndex
    Next chtSeries
End Sub

