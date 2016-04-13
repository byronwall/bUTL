Attribute VB_Name = "Chart_Axes"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : Chart_Axes
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains code related to chart axes
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : Chart_Axis_AutoX
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Reverts the x axis of a chart back to Auto
'---------------------------------------------------------------------------------------
'
Sub Chart_Axis_AutoX()

    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        Dim myChart As Chart

        Dim xAxis As Axis

        Set myChart = chartObj.Chart

        Set xAxis = myChart.Axes(xlCategory)
        xAxis.MaximumScaleIsAuto = True
        xAxis.MinimumScaleIsAuto = True
        xAxis.MajorUnitIsAuto = True
        xAxis.MinorUnitIsAuto = True

    Next chartObj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_Axis_AutoY
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Reverts the Y axis of a chart back to Auto
'---------------------------------------------------------------------------------------
'
Sub Chart_Axis_AutoY()

    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        Dim myChart As Chart

        Dim yAxis As Axis

        Set myChart = chartObj.Chart

        Set yAxis = myChart.Axes(xlValue)
        yAxis.MaximumScaleIsAuto = True
        yAxis.MinimumScaleIsAuto = True
        yAxis.MajorUnitIsAuto = True
        yAxis.MinorUnitIsAuto = True

    Next chartObj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_FitAxisToMaxAndMin
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Iterates through all series and sets desired axis to max/min of data
'---------------------------------------------------------------------------------------
'
Sub Chart_FitAxisToMaxAndMin(typeOfAxis As XlAxisType)
    Dim chartObj As ChartObject
    For Each chartObj In Chart_GetObjectsFromObject(Selection)
        '2015 11 09 moved first inside loop so that it works for multiple charts
        Dim first As Boolean
        first = True

        Dim myChart As Chart
        Set myChart = chartObj.Chart

        Dim mySeries As series
        For Each mySeries In myChart.SeriesCollection

            Dim minimumValue As Double
            Dim maximumValue As Double

            If typeOfAxis = xlCategory Then

                minimumValue = Application.Min(mySeries.XValues)
                maximumValue = Application.Max(mySeries.XValues)

            ElseIf typeOfAxis = xlValue Then

                minimumValue = Application.Min(mySeries.Values)
                maximumValue = Application.Max(mySeries.Values)

            End If


            Dim myAxis As Axis
            Set myAxis = myChart.Axes(typeOfAxis)

            Dim newMinimum As Boolean
            Dim newMaximum As Boolean
            
            newMaximum = maximumValue > myAxis.MaximumScale
            newMinimum = minimumValue < myAxis.MinimumScale

            If first Or newMinimum Then
                myAxis.MinimumScale = minimumValue
            End If
            If first Or newMaximum Then
                myAxis.MaximumScale = maximumValue
            End If

            first = False
        Next mySeries
    Next chartObj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_YAxisRangeWithAvgAndStdev
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets a chart's Y axis to a number of standard deviations
' Flags     : not-used
'---------------------------------------------------------------------------------------
'
Public Sub Chart_YAxisRangeWithAvgAndStdev()
    Dim numberStandardDeviations As Double

    numberStandardDeviations = CDbl(InputBox("How many standard deviations to include?"))

    Dim chartObj As ChartObject

    For Each chartObj In Chart_GetObjectsFromObject(Selection)

        Dim mySeries As series
        Set mySeries = chartObj.Chart.SeriesCollection(1)

        Dim averageValue As Double
        Dim standardValue As Double

        averageValue = WorksheetFunction.Average(mySeries.Values)
        standardValue = WorksheetFunction.StDev(mySeries.Values)

        chartObj.Chart.Axes(xlValue).MinimumScale = averageValue - standardValue * numberStandardDeviations
        chartObj.Chart.Axes(xlValue).MaximumScale = averageValue + standardValue * numberStandardDeviations

    Next

End Sub
