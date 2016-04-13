Attribute VB_Name = "Chart_Format"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : Chart_Format
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains code related to formatting charts
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : Chart_AddTitles
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds all missing titles to all selected charts
'---------------------------------------------------------------------------------------
'
Sub Chart_AddTitles()
    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        If Not myChartObject.Chart.Axes(xlCategory).HasTitle Then
            myChartObject.Chart.Axes(xlCategory).HasTitle = True
            myChartObject.Chart.Axes(xlCategory).AxisTitle.Text = "x axis"
        End If

        If Not myChartObject.Chart.Axes(xlValue).HasTitle Then
            myChartObject.Chart.Axes(xlValue).HasTitle = True
            myChartObject.Chart.Axes(xlValue).AxisTitle.Text = "y axis"
        End If

        If Not myChartObject.Chart.HasTitle Then
            myChartObject.Chart.HasTitle = True
            myChartObject.Chart.ChartTitle.Text = "chart"
        End If

    Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_ApplyFormattingToSelected
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Applies a semi-random format to all charts
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub Chart_ApplyFormattingToSelected()

    Dim myChart As ChartObject

    For Each myChart In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series

        For Each mySeries In myChart.Chart.SeriesCollection
            mySeries.MarkerSize = 5
        Next mySeries
    Next myChart

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_ApplyTrendColors
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Applies the predetermined chart colors to each series
'---------------------------------------------------------------------------------------
'
Sub Chart_ApplyTrendColors()

    Dim myChartObject As ChartObject
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series
        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim ButlSeries As New bUTLChartSeries
            ButlSeries.UpdateFromChartSeries mySeries

            mySeries.MarkerForegroundColorIndex = xlColorIndexNone
            mySeries.MarkerBackgroundColor = Chart_GetColor(ButlSeries.SeriesNumber)

            mySeries.Format.Line.ForeColor.RGB = mySeries.MarkerBackgroundColor

        Next mySeries
    Next myChartObject
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_AxisTitleIsSeriesTitle
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets the y axis title equal to the series name of the last series
'---------------------------------------------------------------------------------------
'
Sub Chart_AxisTitleIsSeriesTitle()

    Dim myChartObject As ChartObject
    Dim myChart As Chart
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
        Set myChart = myChartObject.Chart

        Dim ButlSeries As bUTLChartSeries
        Dim mySeries As series

        For Each mySeries In myChart.SeriesCollection
            Set ButlSeries = New bUTLChartSeries
            ButlSeries.UpdateFromChartSeries mySeries

            myChart.Axes(xlValue, mySeries.AxisGroup).HasTitle = True
            myChart.Axes(xlValue, mySeries.AxisGroup).AxisTitle.Text = ButlSeries.name
            
            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            myChart.Axes(xlCategory).HasTitle = True
            myChart.Axes(xlCategory).AxisTitle.Text = ButlSeries.XValues.Cells(1, 1).Offset(-1).Value

        Next mySeries
    Next myChartObject
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_CreateDataLabels
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds a data label for each series in the chart
' Flag      : new-feature
'---------------------------------------------------------------------------------------
'
Sub Chart_CreateDataLabels()

    Dim myChartObject As ChartObject
    On Error GoTo Chart_CreateDataLabels_Error

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series
        For Each mySeries In myChartObject.Chart.SeriesCollection

            Dim myPoint As Point
            Set myPoint = mySeries.Points(2)

            myPoint.HasDataLabel = False
            myPoint.DataLabel.Position = xlLabelPositionRight
            myPoint.DataLabel.ShowSeriesName = True
            myPoint.DataLabel.ShowValue = False
            myPoint.DataLabel.ShowCategoryName = False
            myPoint.DataLabel.ShowLegendKey = True

        Next mySeries
    Next myChartObject

    On Error GoTo 0
    Exit Sub

Chart_CreateDataLabels_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Chart_CreateDataLabels of Module Chart_Format"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Chart_GridOfCharts
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates a grid of charts.  Used by the form.
'---------------------------------------------------------------------------------------
'
Sub Chart_GridOfCharts( _
    Optional chartColumns As Long = 3, _
    Optional chartWidth As Double = 400, _
    Optional chartHeight As Double = 300, _
    Optional verticalDisplacement As Double = 80, _
    Optional horizontalDisplacement As Double = 40, _
    Optional checkDown As Boolean = False, _
    Optional isZoom As Boolean = False)

    Dim myChartObject As ChartObject

    Dim mySheet As Worksheet
    Set mySheet = ActiveSheet

    Application.ScreenUpdating = False

    Dim count As Long
    count = 0

    For Each myChartObject In mySheet.ChartObjects
        Dim leftSide As Double, topSide As Double

        If checkDown Then
            leftSide = (count \ chartColumns) * chartWidth + horizontalDisplacement
            topSide = (count Mod chartColumns) * chartHeight + verticalDisplacement
        Else
            leftSide = (count Mod chartColumns) * chartWidth + horizontalDisplacement
            topSide = (count \ chartColumns) * chartHeight + verticalDisplacement
        End If

        myChartObject.top = topSide
        myChartObject.left = leftSide
        myChartObject.Width = chartWidth
        myChartObject.Height = chartHeight

        count = count + 1

    Next myChartObject

    'loop through columsn to find how far to zoom
    If isZoom Then
        Dim ColumnZoom As Long
        ColumnZoom = 1
        Do While mySheet.Cells(1, ColumnZoom).left < chartColumns * chartWidth
            ColumnZoom = ColumnZoom + 1
        Loop

        mySheet.Range("A:A", mySheet.Cells(1, ColumnZoom - 1).EntireColumn).Select
        ActiveWindow.Zoom = True
        mySheet.Range("A1").Select
    End If

    Application.ScreenUpdating = True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartApplyToAll
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Forces all charts to be a XYScatter type
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub ChartApplyToAll()

    Dim myChartObject As ChartObject
    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
        myChartObject.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next myChartObject

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartCreateXYGrid
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Creates a matrix of charts similar to pairs in R
'---------------------------------------------------------------------------------------
'
Sub ChartCreateXYGrid()

    On Error GoTo ChartCreateXYGrid_Error

    DeleteAllCharts

    'dataRange will contain the block of data with titles included

    Dim dataRange As Range
    Set dataRange = Application.InputBox("Select data with titles", Type:=8)

    Application.ScreenUpdating = False

    Dim iRow As Long, iCol As Long
    iRow = 0

    Dim myHeight As Double, myWidth As Double
    myHeight = 300
    myWidth = 400

    Dim xColumnData As Range, yColumnData As Range
    For Each yColumnData In dataRange.Columns
        iCol = 0

        For Each xColumnData In dataRange.Columns
            If iRow <> iCol Then
                Dim newChart As Chart
                Set newChart = ActiveSheet.ChartObjects.Add(iCol * myWidth, _
                                                       iRow * myHeight + 100, _
                                                       myWidth, _
                                                       myHeight).Chart

                Dim mySeries As series
                Dim ButlSeries As New bUTLChartSeries

                'offset allows for the title to be excluded
                Set ButlSeries.XValues = Intersect(xColumnData, xColumnData.Offset(1))
                Set ButlSeries.Values = Intersect(yColumnData, yColumnData.Offset(1))
                Set ButlSeries.name = yColumnData.Cells(1)
                ButlSeries.ChartType = xlXYScatter

                Set mySeries = ButlSeries.AddSeriesToChart(newChart)

                mySeries.MarkerSize = 3
                mySeries.MarkerStyle = xlMarkerStyleCircle

                Dim newAxis As Axis
                Set newAxis = newChart.Axes(xlCategory)
                newAxis.HasTitle = True
                newAxis.AxisTitle.Text = xColumnData.Cells(1)
                newAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)
                newAxis.MinorGridlines.Border.Color = RGB(220, 220, 220)

                Set newAxis = newChart.Axes(xlValue)
                newAxis.HasTitle = True
                newAxis.AxisTitle.Text = yColumnData.Cells(1)
                newAxis.MajorGridlines.Border.Color = RGB(200, 200, 200)
                newAxis.MinorGridlines.Border.Color = RGB(220, 220, 220)

                newChart.HasTitle = True
                newChart.ChartTitle.Text = yColumnData.Cells(1) & " vs. " & xColumnData.Cells(1)
                'newChart.ChartTitle.Characters.Font.Size = 8
                newChart.Legend.Delete
            End If

            iCol = iCol + 1
        Next

        iRow = iRow + 1
    Next

    Application.ScreenUpdating = True

    dataRange.Cells(1, 1).Activate

    On Error GoTo 0
    Exit Sub

ChartCreateXYGrid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure ChartCreateXYGrid of Module Chart_Format"
    MsgBox "This is most likely due to Range issues"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartDefaultFormat
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Set the default format for all charts on ActiveSheet
'---------------------------------------------------------------------------------------
'
Sub ChartDefaultFormat()

    Dim myChartObject As ChartObject

    For Each myChartObject In Chart_GetObjectsFromObject(Selection)
        Dim newChart As Chart

        Set newChart = myChartObject.Chart

        Dim mySeries As series
        For Each mySeries In newChart.SeriesCollection

            mySeries.MarkerSize = 3
            mySeries.MarkerStyle = xlMarkerStyleCircle

            If mySeries.ChartType = xlXYScatterLines Then
                mySeries.Format.Line.Weight = 1.5

            End If

            mySeries.MarkerForegroundColorIndex = xlColorIndexNone
            mySeries.MarkerBackgroundColorIndex = xlColorIndexAutomatic

        Next mySeries


        newChart.HasLegend = True
        newChart.Legend.Position = xlLegendPositionBottom

        Dim myAxis As Axis
        Set myAxis = newChart.Axes(xlValue)

        myAxis.MajorGridlines.Border.Color = RGB(242, 242, 242)
        myAxis.Crosses = xlAxisCrossesMinimum
        
        Set myAxis = newChart.Axes(xlCategory)
        
        myAxis.HasMajorGridlines = True

        myAxis.MajorGridlines.Border.Color = RGB(242, 242, 242)

        If newChart.HasTitle Then
            newChart.ChartTitle.Characters.Font.Size = 12
            newChart.ChartTitle.Characters.Font.Bold = True
        End If

        Set myAxis = newChart.Axes(xlCategory)

    Next myChartObject

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartPropMove
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets the "move or size" setting for all charts
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub ChartPropMove()

    Dim chartObj As ChartObject

    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        chartObj.Placement = xlFreeFloating
    Next chartObj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChartTitleEqualsSeriesSelection
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets the chart title equal to the name of the first series
' Flag      : not-used
'---------------------------------------------------------------------------------------
'
Sub ChartTitleEqualsSeriesSelection()

    Dim myChartObject As ChartObject


    For Each myChartObject In Selection
        myChartObject.Chart.ChartTitle.Text = myChartObject.Chart.SeriesCollection(1).name
    Next myChartObject


End Sub

