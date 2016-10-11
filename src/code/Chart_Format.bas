Attribute VB_Name = "Chart_Format"
Option Explicit


Sub Chart_AddTitles()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AddTitles
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds all missing titles to all selected charts
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        If Not chtObj.Chart.Axes(xlCategory).HasTitle Then
            chtObj.Chart.Axes(xlCategory).HasTitle = True
            chtObj.Chart.Axes(xlCategory).AxisTitle.Text = "x axis"
        End If

        If Not chtObj.Chart.Axes(xlValue, xlPrimary).HasTitle Then
            chtObj.Chart.Axes(xlValue).HasTitle = True
            chtObj.Chart.Axes(xlValue).AxisTitle.Text = "y axis"
        End If

        '2015 12 14, add support for 2nd y axis
        If chtObj.Chart.Axes.count = 3 Then
            If Not chtObj.Chart.Axes(xlValue, xlSecondary).HasTitle Then
                chtObj.Chart.Axes(xlValue, xlSecondary).HasTitle = True
                chtObj.Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = "2nd y axis"
            End If
        End If

        If Not chtObj.Chart.HasTitle Then
            chtObj.Chart.HasTitle = True
            chtObj.Chart.ChartTitle.Text = "chart"
        End If

    Next

End Sub


Sub Chart_ApplyFormattingToSelected()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyFormattingToSelected
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies a semi-random format to all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series

        For Each chtSeries In chtObj.Chart.SeriesCollection
            chtSeries.MarkerSize = 5
        Next chtSeries
    Next chtObj

End Sub


Sub Chart_ApplyTrendColors()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyTrendColors
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies the predetermined chart colors to each series
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series
        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries chtSeries

            chtSeries.MarkerForegroundColorIndex = xlColorIndexNone
            chtSeries.MarkerBackgroundColor = Chart_GetColor(butlSeries.SeriesNumber)

            chtSeries.Format.Line.ForeColor.RGB = chtSeries.MarkerBackgroundColor

        Next chtSeries
    Next chtObj
End Sub


Sub Chart_AxisTitleIsSeriesTitle()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_AxisTitleIsSeriesTitle
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the y axis title equal to the series name of the last series
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    Dim cht As Chart
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        Set cht = chtObj.Chart

        Dim butlSeries As bUTLChartSeries
        Dim chtSeries As series

        For Each chtSeries In cht.SeriesCollection
            Set butlSeries = New bUTLChartSeries
            butlSeries.UpdateFromChartSeries chtSeries

            cht.Axes(xlValue, chtSeries.AxisGroup).HasTitle = True
            cht.Axes(xlValue, chtSeries.AxisGroup).AxisTitle.Text = butlSeries.name

            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            '2015 12 14, add a check to ensure that the XValue exists
            If Not butlSeries.XValues Is Nothing Then
                cht.Axes(xlCategory).HasTitle = True
                cht.Axes(xlCategory).AxisTitle.Text = butlSeries.XValues.Cells(1, 1).Offset(-1).Value
            End If

        Next chtSeries
    Next chtObj
End Sub


Sub Chart_CreateDataLabels()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_CreateDataLabels
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds a data label for each series in the chart
    ' Flag      : new-feature
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    On Error GoTo Chart_CreateDataLabels_Error

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series
        For Each chtSeries In chtObj.Chart.SeriesCollection

            Dim dataPoint As Point
            Set dataPoint = chtSeries.Points(2)

            dataPoint.HasDataLabel = False
            dataPoint.DataLabel.Position = xlLabelPositionRight
            dataPoint.DataLabel.ShowSeriesName = True
            dataPoint.DataLabel.ShowValue = False
            dataPoint.DataLabel.ShowCategoryName = False
            dataPoint.DataLabel.ShowLegendKey = True

        Next chtSeries
    Next chtObj

    On Error GoTo 0
    Exit Sub

Chart_CreateDataLabels_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Chart_CreateDataLabels of Module Chart_Format"

End Sub



Sub Chart_GridOfCharts( _
    Optional columnCount As Long = 3, _
    Optional chtWidth As Double = 400, _
    Optional chtHeight As Double = 300, _
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
    Dim chtObj As ChartObject

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Application.ScreenUpdating = False

    Dim chtCount As Long
    chtCount = 0

    For Each chtObj In sht.ChartObjects
        Dim left As Double, top As Double

        If shouldFillDownFirst Then
            left = (chtCount \ columnCount) * chtWidth + offsetHorizontal
            top = (chtCount Mod columnCount) * chtHeight + offsetVertical
        Else
            left = (chtCount Mod columnCount) * chtWidth + offsetHorizontal
            top = (chtCount \ columnCount) * chtHeight + offsetVertical
        End If

        chtObj.top = top
        chtObj.left = left
        chtObj.Width = chtWidth
        chtObj.Height = chtHeight

        chtCount = chtCount + 1

    Next chtObj

    'loop through columsn to find how far to zoom
    If shouldZoomOnGrid Then
        Dim columnToZoomTo As Long
        columnToZoomTo = 1
        Do While sht.Cells(1, columnToZoomTo).left < columnCount * chtWidth
            columnToZoomTo = columnToZoomTo + 1
        Loop

        sht.Range("A:A", sht.Cells(1, columnToZoomTo - 1).EntireColumn).Select
        ActiveWindow.Zoom = True
        sht.Range("A1").Select
    End If

    Application.ScreenUpdating = True

End Sub


Sub ChartApplyToAll()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartApplyToAll
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces all charts to be a XYScatter type
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject
    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        chtObj.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next chtObj

End Sub


Sub ChartCreateXYGrid()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartCreateXYGrid
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Creates a matrix of charts similar to pairs in R
    '---------------------------------------------------------------------------------------
    '
    On Error GoTo ChartCreateXYGrid_Error

    DeleteAllCharts

    'rng_data will contain the block of data with titles included

    Dim rngData As Range
    Set rngData = Application.InputBox("Select data with titles", Type:=8)

    Application.ScreenUpdating = False

    Dim rowIndex As Long, columnIndex As Long
    rowIndex = 0

    Dim chtHeight As Double, chtWidth As Double
    chtHeight = 300
    chtWidth = 400

    Dim rngColXData As Range, rngColYData As Range
    For Each rngColYData In rngData.Columns
        columnIndex = 0

        For Each rngColXData In rngData.Columns
            If rowIndex <> columnIndex Then
                Dim cht As Chart
                Set cht = ActiveSheet.ChartObjects.Add(columnIndex * chtWidth, _
                                                       rowIndex * chtHeight + 100, _
                                                       chtWidth, _
                                                       chtHeight).Chart

                Dim chtSeries As series
                Dim butlSeries As New bUTLChartSeries

                'offset allows for the title to be excluded
                Set butlSeries.XValues = Intersect(rngColXData, rngColXData.Offset(1))
                Set butlSeries.Values = Intersect(rngColYData, rngColYData.Offset(1))
                Set butlSeries.name = rngColYData.Cells(1)
                butlSeries.ChartType = xlXYScatter

                Set chtSeries = butlSeries.AddSeriesToChart(cht)

                chtSeries.MarkerSize = 3
                chtSeries.MarkerStyle = xlMarkerStyleCircle

                Dim chtAxis As Axis
                Set chtAxis = cht.Axes(xlCategory)
                chtAxis.HasTitle = True
                chtAxis.AxisTitle.Text = rngColXData.Cells(1)
                chtAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)
                chtAxis.MinorGridlines.Border.Color = RGB(220, 220, 220)

                Set chtAxis = cht.Axes(xlValue)
                chtAxis.HasTitle = True
                chtAxis.AxisTitle.Text = rngColYData.Cells(1)
                chtAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)
                chtAxis.MinorGridlines.Border.Color = RGB(220, 220, 220)

                cht.HasTitle = True
                cht.ChartTitle.Text = rngColYData.Cells(1) & " vs. " & rngColXData.Cells(1)
                'cht.ChartTitle.Characters.Font.Size = 8
                cht.Legend.Delete
            End If

            columnIndex = columnIndex + 1
        Next

        rowIndex = rowIndex + 1
    Next

    Application.ScreenUpdating = True

    rngData.Cells(1, 1).Activate

    On Error GoTo 0
    Exit Sub

ChartCreateXYGrid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure ChartCreateXYGrid of Module Chart_Format"
    MsgBox "This is most likely due to Range issues"

End Sub


Sub ChartDefaultFormat()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartDefaultFormat
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Set the default format for all charts on ActiveSheet
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart

        Set cht = chtObj.Chart

        Dim chtSeries As series
        For Each chtSeries In cht.SeriesCollection

            chtSeries.MarkerSize = 3
            chtSeries.MarkerStyle = xlMarkerStyleCircle

            If chtSeries.ChartType = xlXYScatterLines Then
                chtSeries.Format.Line.Weight = 1.5

            End If

            chtSeries.MarkerForegroundColorIndex = xlColorIndexNone
            chtSeries.MarkerBackgroundColorIndex = xlColorIndexAutomatic

        Next chtSeries


        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionBottom

        Dim chtAxis As Axis
        Set chtAxis = cht.Axes(xlValue)

        chtAxis.MajorGridlines.Border.Color = RGB(242, 242, 242)
        chtAxis.Crosses = xlAxisCrossesMinimum
        
        Set chtAxis = cht.Axes(xlCategory)
        
        chtAxis.HasMajorGridlines = True

        chtAxis.MajorGridlines.Border.Color = RGB(242, 242, 242)

        If cht.HasTitle Then
            cht.ChartTitle.Characters.Font.Size = 12
            cht.ChartTitle.Characters.Font.Bold = True
        End If

        Set chtAxis = cht.Axes(xlCategory)

    Next chtObj

End Sub


Sub ChartPropMove()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartPropMove
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the "move or size" setting for all charts
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Chart_GetObjectsFromObject(Selection)
        chtObj.Placement = xlFreeFloating
    Next chtObj

End Sub


Sub ChartTitleEqualsSeriesSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : ChartTitleEqualsSeriesSelection
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets the chart title equal to the name of the first series
    ' Flag      : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim chtObj As ChartObject

    For Each chtObj In Selection
        chtObj.Chart.ChartTitle.Text = chtObj.Chart.SeriesCollection(1).name
    Next chtObj
    
End Sub

