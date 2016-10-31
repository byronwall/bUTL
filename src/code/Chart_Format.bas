Attribute VB_Name = "Chart_Format"
Option Explicit


Public Sub Chart_AddTitles()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AddTitles
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds all missing titles to all selected charts
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Const X_AXIS_TITLE As String = "x axis"
    Const Y_AXIS_TITLE As String = "y axis"
    Const SECOND_Y_AXIS_TITLE As String = "2nd y axis"
    Const CHART_TITLE As String = "chart"

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        With targetObject.Chart
            If Not .Axes(xlCategory).HasTitle Then
                .Axes(xlCategory).HasTitle = True
                .Axes(xlCategory).AxisTitle.Text = X_AXIS_TITLE
            End If
    
            If Not .Axes(xlValue, xlPrimary).HasTitle Then
                .Axes(xlValue).HasTitle = True
                .Axes(xlValue).AxisTitle.Text = Y_AXIS_TITLE
            End If
    
            '2015 12 14, add support for 2nd y axis
            If .Axes.count = 3 Then
                If Not .Axes(xlValue, xlSecondary).HasTitle Then
                    .Axes(xlValue, xlSecondary).HasTitle = True
                    .Axes(xlValue, xlSecondary).AxisTitle.Text = SECOND_Y_AXIS_TITLE
                End If
            End If
    
            If Not .HasTitle Then
                .HasTitle = True
                .ChartTitle.Text = CHART_TITLE
            End If
        End With
    Next targetObject

End Sub


Public Sub Chart_ApplyFormattingToSelected()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyFormattingToSelected
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies a semi-random format to all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Const MARKER_SIZE As Long = 5

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series

        For Each targetSeries In targetObject.Chart.SeriesCollection
            targetSeries.MarkerSize = MARKER_SIZE
        Next targetSeries
    Next targetObject

End Sub


Public Sub Chart_ApplyTrendColors()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyTrendColors
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies the predetermined chart colors to each series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerBackgroundColor = Chart_GetColor(butlSeries.SeriesNumber)

            targetSeries.Format.Line.ForeColor.RGB = targetSeries.MarkerBackgroundColor

        Next targetSeries
    Next targetObject
End Sub


Public Sub Chart_AxisTitleIsSeriesTitle()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AxisTitleIsSeriesTitle
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the y axis title equal to the series name of the last series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    Dim targetChart As Chart
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Set targetChart = targetObject.Chart

        Dim butlSeries As bUTLChartSeries
        Dim targetSeries As series

        For Each targetSeries In targetChart.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetChart.Axes(xlValue, targetSeries.AxisGroup).HasTitle = True
            targetChart.Axes(xlValue, targetSeries.AxisGroup).AxisTitle.Text = butlSeries.name

            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            '2015 12 14, add a check to ensure that the XValue exists
            If Not butlSeries.XValues Is Nothing Then
                targetChart.Axes(xlCategory).HasTitle = True
                targetChart.Axes(xlCategory).AxisTitle.Text = butlSeries.XValues.Cells(1, 1).Offset(-1).Value
            End If

        Next targetSeries
    Next targetObject
End Sub


Public Sub Chart_CreateDataLabels()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_CreateDataLabels
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds a data label for each series in the chart
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    On Error GoTo Chart_CreateDataLabels_Error

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim dataPoint As Point
            Set dataPoint = targetSeries.Points(2)

            dataPoint.HasDataLabel = False
            dataPoint.DataLabel.Position = xlLabelPositionRight
            dataPoint.DataLabel.ShowSeriesName = True
            dataPoint.DataLabel.ShowValue = False
            dataPoint.DataLabel.ShowCategoryName = False
            dataPoint.DataLabel.ShowLegendKey = True

        Next targetSeries
    Next targetObject

    On Error GoTo 0
    Exit Sub

Chart_CreateDataLabels_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Chart_CreateDataLabels of Module Chart_Format"

End Sub



Public Sub Chart_GridOfCharts( _
    Optional columnCount As Long = 3, _
    Optional chartWidth As Double = 400, _
    Optional chartHeight As Double = 300, _
    Optional offsetVertical As Double = 80, _
    Optional offsetHorizontal As Double = 40, _
    Optional shouldFillDownFirst As Boolean = False, _
    Optional shouldZoomOnGrid As Boolean = False)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GridOfCharts
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a grid of charts.  Used by the form.
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet

    Application.ScreenUpdating = False

    Dim countOfCharts As Long
    countOfCharts = 0

    For Each targetObject In targetSheet.ChartObjects
        Dim left As Double, top As Double

        If shouldFillDownFirst Then
            left = (countOfCharts \ columnCount) * chartWidth + offsetHorizontal
            top = (countOfCharts Mod columnCount) * chartHeight + offsetVertical
        Else
            left = (countOfCharts Mod columnCount) * chartWidth + offsetHorizontal
            top = (countOfCharts \ columnCount) * chartHeight + offsetVertical
        End If

        targetObject.top = top
        targetObject.left = left
        targetObject.Width = chartWidth
        targetObject.Height = chartHeight

        countOfCharts = countOfCharts + 1

    Next targetObject

    'loop through columns to find how far to zoom
    'Cells.Left property returns a variant in points
    If shouldZoomOnGrid Then
        Dim columnToZoomTo As Long
        columnToZoomTo = 1
        Do While targetSheet.Cells(1, columnToZoomTo).left < columnCount * chartWidth
            columnToZoomTo = columnToZoomTo + 1
        Loop

        targetSheet.Range("A:A", targetSheet.Cells(1, columnToZoomTo - 1).EntireColumn).Select
        ActiveWindow.Zoom = True
        targetSheet.Range("A1").Select
    End If

    Application.ScreenUpdating = True

End Sub


Public Sub ChartApplyToAll()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartApplyToAll
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces all charts to be a XYScatter type
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next targetObject

End Sub


Public Sub ChartCreateXYGrid()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartCreateXYGrid
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Creates a matrix of charts similar to pairs in R
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo ChartCreateXYGrid_Error

    DeleteAllCharts
    'VBA doesn't allow a constant to be defined using a function (rgb) so we use a local variable rather than
    'muddying it up with the calculated value of the rgb function
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(200, 200, 200)
    Dim minorGridlineColor As Long
    minorGridlineColor = RGB(220, 220, 220)
    
    Const CHART_HEIGHT As Long = 300
    Const CHART_WIDTH As Long = 400
    Const MARKER_SIZE As Long = 3
    'dataRange will contain the block of data with titles included
    Dim dataRange As Range
    Set dataRange = Application.InputBox("Select data with titles", Type:=8)

    Application.ScreenUpdating = False

    Dim rowIndex As Long, columnIndex As Long
    rowIndex = 0

    Dim xAxisDataRange As Range, yAxisDataRange As Range
    For Each yAxisDataRange In dataRange.Columns
        columnIndex = 0

        For Each xAxisDataRange In dataRange.Columns
            If rowIndex <> columnIndex Then
                Dim targetChart As Chart
                Set targetChart = ActiveSheet.ChartObjects.Add(columnIndex * CHART_WIDTH, _
                                                               rowIndex * CHART_HEIGHT + 100, _
                                                               CHART_WIDTH, CHART_HEIGHT).Chart

                Dim targetSeries As series
                Dim butlSeries As New bUTLChartSeries

                'offset allows for the title to be excluded
                Set butlSeries.XValues = Intersect(xAxisDataRange, xAxisDataRange.Offset(1))
                Set butlSeries.Values = Intersect(yAxisDataRange, yAxisDataRange.Offset(1))
                Set butlSeries.name = yAxisDataRange.Cells(1)
                butlSeries.ChartType = xlXYScatter

                Set targetSeries = butlSeries.AddSeriesToChart(targetChart)

                targetSeries.MarkerSize = MARKER_SIZE
                targetSeries.MarkerStyle = xlMarkerStyleCircle

                Dim targetAxis As Axis
                Set targetAxis = targetChart.Axes(xlCategory)
                targetAxis.HasTitle = True
                targetAxis.AxisTitle.Text = xAxisDataRange.Cells(1)
                targetAxis.MajorGridlines.Border.Color = majorGridlineColor
                targetAxis.MinorGridlines.Border.Color = minorGridlineColor

                Set targetAxis = targetChart.Axes(xlValue)
                targetAxis.HasTitle = True
                targetAxis.AxisTitle.Text = yAxisDataRange.Cells(1)
                targetAxis.MajorGridlines.Border.Color = majorGridlineColor
                targetAxis.MinorGridlines.Border.Color = minorGridlineColor

                targetChart.HasTitle = True
                targetChart.ChartTitle.Text = yAxisDataRange.Cells(1) & " vs. " & xAxisDataRange.Cells(1)
                'targetChart.ChartTitle.Characters.Font.Size = 8
                targetChart.Legend.Delete
            End If

            columnIndex = columnIndex + 1
        Next xAxisDataRange

        rowIndex = rowIndex + 1
    Next yAxisDataRange

    Application.ScreenUpdating = True

    dataRange.Cells(1, 1).Activate

    On Error GoTo 0
    Exit Sub

ChartCreateXYGrid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure ChartCreateXYGrid of Module Chart_Format"
    MsgBox "This is most likely due to Range issues"

End Sub


Public Sub ChartDefaultFormat()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartDefaultFormat
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Set the default format for all charts on ActiveSheet
    '---------------------------------------------------------------------------------------
    '
    Const MARKER_SIZE As Long = 3
    Dim majorGridlineColor As Long
    majorGridlineColor = RGB(242, 242, 242)
    Const TITLE_FONT_SIZE As Long = 12
    Const SERIES_LINE_WEIGHT As Single = 1.5
    
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart

        Set targetChart = targetObject.Chart

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection

            targetSeries.MarkerSize = MARKER_SIZE
            targetSeries.MarkerStyle = xlMarkerStyleCircle

            If targetSeries.ChartType = xlXYScatterLines Then targetSeries.Format.Line.Weight = SERIES_LINE_WEIGHT

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerBackgroundColorIndex = xlColorIndexAutomatic

        Next targetSeries


        targetChart.HasLegend = True
        targetChart.Legend.Position = xlLegendPositionBottom

        Dim targetAxis As Axis
        Set targetAxis = targetChart.Axes(xlValue)

        targetAxis.MajorGridlines.Border.Color = majorGridlineColor
        targetAxis.Crosses = xlAxisCrossesMinimum
        
        Set targetAxis = targetChart.Axes(xlCategory)
        
        targetAxis.HasMajorGridlines = True

        targetAxis.MajorGridlines.Border.Color = majorGridlineColor

        If targetChart.HasTitle Then
            targetChart.ChartTitle.Characters.Font.Size = TITLE_FONT_SIZE
            targetChart.ChartTitle.Characters.Font.Bold = True
        End If

        Set targetAxis = targetChart.Axes(xlCategory)

    Next targetObject

End Sub


Public Sub ChartPropMove()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartPropMove
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the "move or size" setting for all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        targetObject.Placement = xlFreeFloating
    Next targetObject

End Sub


Public Sub ChartTitleEqualsSeriesSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartTitleEqualsSeriesSelection
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the chart title equal to the name of the first series
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject

    For Each targetObject In Selection
        targetObject.Chart.ChartTitle.Text = targetObject.Chart.SeriesCollection(1).name
    Next targetObject
    
End Sub

