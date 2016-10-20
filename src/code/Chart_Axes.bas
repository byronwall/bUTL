Attribute VB_Name = "Chart_Axes"
Option Explicit

Public Sub Chart_Axis_AutoX()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoX
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the x axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        
        Dim xAxis As Axis
        Set xAxis = targetChart.Axes(xlCategory)
        xAxis.MaximumScaleIsAuto = True
        xAxis.MinimumScaleIsAuto = True
        xAxis.MajorUnitIsAuto = True
        xAxis.MinorUnitIsAuto = True

    Next targetObject

End Sub


Public Sub Chart_Axis_AutoY()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_Axis_AutoY
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Reverts the Y axis of a chart back to Auto
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        Dim targetChart As Chart
        Set targetChart = targetObject.Chart
        
        Dim yAxis As Axis
        Set yAxis = targetChart.Axes(xlValue)
        yAxis.MaximumScaleIsAuto = True
        yAxis.MinimumScaleIsAuto = True
        yAxis.MajorUnitIsAuto = True
        yAxis.MinorUnitIsAuto = True

    Next targetObject

End Sub


Public Sub Chart_FitAxisToMaxAndMin(ByVal axisType As XlAxisType)
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_FitAxisToMaxAndMin
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Iterates through all series and sets desired axis to max/min of data
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)
        '2015 11 09 moved first inside loop so that it works for multiple charts
        Dim isFirst As Boolean
        isFirst = True

        Dim targetChart As Chart
        Set targetChart = targetObject.Chart

        Dim targetSeries As series
        For Each targetSeries In targetChart.SeriesCollection

            Dim minSeriesValue As Double
            Dim maxSeriesValue As Double

            If axisType = xlCategory Then

                minSeriesValue = Application.Min(targetSeries.XValues)
                maxSeriesValue = Application.Max(targetSeries.XValues)

            ElseIf axisType = xlValue Then

                minSeriesValue = Application.Min(targetSeries.Values)
                maxSeriesValue = Application.Max(targetSeries.Values)

            End If

            Dim targetAxis As Axis
            Set targetAxis = targetChart.Axes(axisType)

            Dim isNewMax As Boolean, isNewMin As Boolean
            isNewMax = maxSeriesValue > targetAxis.MaximumScale
            isNewMin = minSeriesValue < targetAxis.MinimumScale

            If isFirst Or isNewMin Then targetAxis.MinimumScale = minSeriesValue
            If isFirst Or isNewMax Then targetAxis.MaximumScale = maxSeriesValue

            isFirst = False
        Next targetSeries
    Next targetObject

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

    Dim targetObject As ChartObject

    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        Set targetSeries = targetObject.Chart.SeriesCollection(1)

        Dim avgSeriesValue As Double
        Dim stdSeriesValue As Double

        avgSeriesValue = WorksheetFunction.Average(targetSeries.Values)
        stdSeriesValue = WorksheetFunction.StDev(targetSeries.Values)

        targetObject.Chart.Axes(xlValue).MinimumScale = avgSeriesValue - stdSeriesValue * numberOfStdDevs
        targetObject.Chart.Axes(xlValue).MaximumScale = avgSeriesValue + stdSeriesValue * numberOfStdDevs

    Next

End Sub
