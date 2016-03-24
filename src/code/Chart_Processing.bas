Attribute VB_Name = "Chart_Processing"
'---------------------------------------------------------------------------------------
' Module    : Chart_Processing
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains some of the heavy lifting processing code for charts
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Chart_CreateChartWithSeriesForEachColumn()
'will create a chart that includes a series with no x value for each column

    Dim dataRange As Range
    Set dataRange = GetInputOrSelection("Select chart data")

    'create a chart
    Dim myChart As ChartObject
    Set myChart = ActiveSheet.ChartObjects.Add(0, 0, 300, 300)
    
    myChart.Chart.ChartType = xlXYScatter

    Dim rangeColumn As Range
    For Each rangeColumn In dataRange.Columns

        Dim chartRange As Range
        Set chartRange = RangeEnd(rangeColumn.Cells(1, 1), xlDown)
        
        Dim ButlSeries As New bUTLChartSeries
        Set ButlSeries.Values = chartRange
        
        ButlSeries.AddSeriesToChart myChart.Chart
    Next

End Sub

Public Sub Chart_CopyToSheet()

    Dim myChart As ChartObject
    
    Dim allObjects As Object
    Set allObjects = Selection
    
    Dim wantNewSheet As VbMsgBoxResult
    wantNewSheet = MsgBox("New sheet?", vbYesNo, "New sheet?")
    
    Dim newSheet As Worksheet
    If wantNewSheet = vbYes Then
        Set newSheet = Worksheets.Add()
    Else
        Set newSheet = Application.InputBox("Pick a cell on a sheet", "Pick sheet", Type:=8).Parent
    End If
    
    For Each myChart In Chart_GetObjectsFromObject(allObjects)
        myChart.Copy
        
        newSheet.Paste
    Next
    
    newSheet.Activate
End Sub

Sub Chart_SortSeriesByName()
'this will sort series by names
    Dim myChart As ChartObject
    For Each myChart In Chart_GetObjectsFromObject(Selection)

        'uses a simple bubble sort but it works... shouldn't have 1000 series anyways
        Dim firstChart As Long
        Dim secondChart As Long
        For firstChart = 1 To myChart.Chart.SeriesCollection.count
            For secondChart = (firstChart + 1) To myChart.Chart.SeriesCollection.count

                Dim FirstButlSeries As New bUTLChartSeries
                Dim SecondButlSeries As New bUTLChartSeries

                FirstButlSeries.UpdateFromChartSeries myChart.Chart.SeriesCollection(firstChart)
                SecondButlSeries.UpdateFromChartSeries myChart.Chart.SeriesCollection(secondChart)

                If FirstButlSeries.name.Value > SecondButlSeries.name.Value Then
                    Dim numberSeries As Long
                    numberSeries = SecondButlSeries.SeriesNumber
                    SecondButlSeries.SeriesNumber = FirstButlSeries.SeriesNumber
                    FirstButlSeries.SeriesNumber = numberSeries

                    SecondButlSeries.UpdateSeriesWithNewValues
                    FirstButlSeries.UpdateSeriesWithNewValues
                End If
            Next
        Next
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_TimeSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper Sub to create a set of charts with the same x axis and varying y
'---------------------------------------------------------------------------------------
'
Sub Chart_TimeSeries(dateRange As Range, dataRange As Range, titles As Range)

    Dim counter As Long
    counter = 1

    Dim title As Range
    Dim rangeColumn As Range

    For Each title In titles

        Dim myChartObject As ChartObject
        Set myChartObject = ActiveSheet.ChartObjects.Add(counter * 300, 0, 300, 300)

        Dim myChart As Chart
        Set myChart = myChartObject.Chart
        myChart.ChartType = xlXYScatterLines
        myChart.HasTitle = True
        myChart.Legend.Delete

        Dim myAxis As Axis
        Set myAxis = myChart.Axes(xlValue)
        myAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)

        Dim mySeries As series
        Dim ButlSeries As New bUTLChartSeries

        Set ButlSeries.XValues = dateRange
        Set ButlSeries.Values = dataRange.Columns(counter)
        Set ButlSeries.name = title

        Set mySeries = ButlSeries.AddSeriesToChart(myChart)

        mySeries.MarkerSize = 3
        mySeries.MarkerStyle = xlMarkerStyleCircle

        counter = counter + 1

    Next title
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

    Dim dateRange As Range
    Dim dataRange As Range
    Dim titles As Range

    'dates are in B4 and down
    Set dateRange = RangeEnd_Boundary(Range("B4"), xlDown)

    'data starts in C4, down and over
    Set dataRange = RangeEnd_Boundary(Range("C4"), xlDown, xlToRight)

    'titels are C2 and over
    Set titles = RangeEnd_Boundary(Range("C2"), xlToRight)

    Chart_TimeSeries dateRange, dataRange, titles
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

    Dim dateRange As Range
    Dim dataRange As Range
    Dim titles As Range

    On Error GoTo CreateMultipleTimeSeries_Error

    DeleteAllCharts

    Set dateRange = Application.InputBox("Select date range", Type:=8)
    Set dataRange = Application.InputBox("Select data", Type:=8)
    Set titles = Application.InputBox("Select titles", Type:=8)

    Chart_TimeSeries dateRange, dataRange, titles

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
    Dim myChart As Chart
    Set myChart = ActiveChart

    Dim mySeries As series
    For Each mySeries In myChart.SeriesCollection

        Dim myValues As Variant
        myValues = mySeries.Values

        'include this line if you want to reestablish labels before deleting
        mySeries.ApplyDataLabels xlDataLabelsShowLabel, , , , True, False, False, False, False

        'loop through values and delete 0-value labels
        Dim i As Long
        For i = LBound(myValues) To UBound(myValues)
            If myValues(i) = 0 Then
                With mySeries.Points(i)
                    If .HasDataLabel Then
                        .DataLabel.Delete
                    End If
                End With
            End If
        Next i
    Next mySeries
End Sub

