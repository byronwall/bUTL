Attribute VB_Name = "Chart_Axes"
Option Explicit

Sub Chart_Axis_AutoX()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoX
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the x axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart
        Set cht = chartObj.Chart
        
        Dim xAxis As Axis
        Set xAxis = cht.Axes(xlCategory)
        xAxis.MaximumScaleIsAuto = True
        xAxis.MinimumScaleIsAuto = True
        xAxis.MajorUnitIsAuto = True
        xAxis.MinorUnitIsAuto = True

    Next chartObj

End Sub


Sub Chart_Axis_AutoY()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoY
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the Y axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart
        Set cht = chartObj.Chart
        
        Dim yAxis As Axis
        Set yAxis = cht.Axes(xlValue)
        yAxis.MaximumScaleIsAuto = True
        yAxis.MinimumScaleIsAuto = True
        yAxis.MajorUnitIsAuto = True
        yAxis.MinorUnitIsAuto = True

    Next chartObj

End Sub


Sub Chart_FitAxisToMaxAndMin(axisType As XlAxisType)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_FitAxisToMaxAndMin
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Iterates through all series and sets desired axis to max/min of data
    '---------------------------------------------------------------------------------------
    '
    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        '2015 11 09 moved first inside loop so that it works for multiple charts
        Dim isFirst As Boolean
        isFirst = True

        Dim cht As Chart
        Set cht = chartObj.Chart

        Dim chtSeries As series
        For Each chtSeries In cht.SeriesCollection

            Dim minValue As Double
            Dim maxValue As Double

            If axisType = xlCategory Then

                minValue = Application.Min(chtSeries.XValues)
                maxValue = Application.Max(chtSeries.XValues)

            ElseIf axisType = xlValue Then

                minValue = Application.Min(chtSeries.Values)
                maxValue = Application.Max(chtSeries.Values)

            End If


            Dim ax As Axis
            Set ax = cht.Axes(axisType)

            Dim isNewMax As Boolean, isNewMin As Boolean
            isNewMax = maxValue > ax.MaximumScale
            isNewMin = minValue < ax.MinimumScale

            If isFirst Or isNewMin Then
                ax.MinimumScale = minValue
            End If
            If isFirst Or isNewMax Then
                ax.MaximumScale = maxValue
            End If

            isFirst = False
        Next chtSeries
    Next chartObj

End Sub


Public Sub Chart_YAxisRangeWithAvgAndStdev()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_YAxisRangeWithAvgAndStdev
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sets a chart's Y axis to a number of standard deviations
    ' Flags     : not-used
    '---------------------------------------------------------------------------------------
    '
    Dim numberOfStdDevs As Double

    numberOfStdDevs = CDbl(InputBox("How many standard deviations to include?"))

    Dim chartObj As ChartObject

    For Each chartObj In Chart_GetObjectsFromObject(Selection)

        Dim chtSeries As series
        Set chtSeries = chartObj.Chart.SeriesCollection(1)

        Dim avgValue As Double
        Dim stdValue As Double

        avgValue = WorksheetFunction.Average(chtSeries.Values)
        stdValue = WorksheetFunction.StDev(chtSeries.Values)

        chartObj.Chart.Axes(xlValue).MinimumScale = avgValue - stdValue * numberOfStdDevs
        chartObj.Chart.Axes(xlValue).MaximumScale = avgValue + stdValue * numberOfStdDevs

    Next

End Sub
