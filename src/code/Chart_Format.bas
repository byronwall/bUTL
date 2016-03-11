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
    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        If Not cht_obj.Chart.Axes(xlCategory).HasTitle Then
            cht_obj.Chart.Axes(xlCategory).HasTitle = True
            cht_obj.Chart.Axes(xlCategory).AxisTitle.Text = "x axis"
        End If

        If Not cht_obj.Chart.Axes(xlValue).HasTitle Then
            cht_obj.Chart.Axes(xlValue).HasTitle = True
            cht_obj.Chart.Axes(xlValue).AxisTitle.Text = "y axis"
        End If

        If Not cht_obj.Chart.HasTitle Then
            cht_obj.Chart.HasTitle = True
            cht_obj.Chart.ChartTitle.Text = "chart"
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

    Dim obj As ChartObject

    For Each obj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series

        For Each ser In obj.Chart.SeriesCollection
            ser.MarkerSize = 5
        Next ser
    Next obj

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_ApplyTrendColors
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Applies the predetermined chart colors to each series
'---------------------------------------------------------------------------------------
'
Sub Chart_ApplyTrendColors()

    Dim cht_obj As ChartObject
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series
        For Each ser In cht_obj.Chart.SeriesCollection

            Dim b_ser As New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            ser.MarkerForegroundColorIndex = xlColorIndexNone
            ser.MarkerBackgroundColor = Chart_GetColor(b_ser.SeriesNumber)

            ser.Format.Line.ForeColor.RGB = ser.MarkerBackgroundColor

        Next ser
    Next cht_obj
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Chart_AxisTitleIsSeriesTitle
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets the y axis title equal to the series name of the last series
'---------------------------------------------------------------------------------------
'
Sub Chart_AxisTitleIsSeriesTitle()

    Dim cht_obj As ChartObject
    Dim cht As Chart
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Set cht = cht_obj.Chart

        Dim b_ser As bUTLChartSeries
        Dim ser As series

        For Each ser In cht.SeriesCollection
            Set b_ser = New bUTLChartSeries
            b_ser.UpdateFromChartSeries ser

            cht.Axes(xlValue, ser.AxisGroup).HasTitle = True
            cht.Axes(xlValue, ser.AxisGroup).AxisTitle.Text = b_ser.name
            
            '2015 11 11, adds the x-title assuming that the name is one cell above the data
            cht.Axes(xlCategory).HasTitle = True
            cht.Axes(xlCategory).AxisTitle.Text = b_ser.XValues.Cells(1, 1).Offset(-1).Value

        Next ser
    Next cht_obj
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

    Dim chtObj As ChartObject
    On Error GoTo Chart_CreateDataLabels_Error

    For Each chtObj In Chart_GetObjectsFromObject(Selection)

        Dim ser As series
        For Each ser In chtObj.Chart.SeriesCollection

            Dim p As Point
            Set p = ser.Points(2)

            p.HasDataLabel = False
            p.DataLabel.Position = xlLabelPositionRight
            p.DataLabel.ShowSeriesName = True
            p.DataLabel.ShowValue = False
            p.DataLabel.ShowCategoryName = False
            p.DataLabel.ShowLegendKey = True

        Next ser
    Next chtObj

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
    Optional int_cols As Long = 3, _
    Optional cht_wid As Double = 400, _
    Optional cht_height As Double = 300, _
    Optional v_off As Double = 80, _
    Optional h_off As Double = 40, _
    Optional chk_down As Boolean = False, _
    Optional bool_zoom As Boolean = False)

    Dim cht_obj As ChartObject

    Dim sht As Worksheet
    Set sht = ActiveSheet

    Application.ScreenUpdating = False

    Dim count As Long
    count = 0

    For Each cht_obj In sht.ChartObjects
        Dim left As Double, top As Double

        If chk_down Then
            left = (count \ int_cols) * cht_wid + h_off
            top = (count Mod int_cols) * cht_height + v_off
        Else
            left = (count Mod int_cols) * cht_wid + h_off
            top = (count \ int_cols) * cht_height + v_off
        End If

        cht_obj.top = top
        cht_obj.left = left
        cht_obj.Width = cht_wid
        cht_obj.Height = cht_height

        count = count + 1

    Next cht_obj

    'loop through columsn to find how far to zoom
    If bool_zoom Then
        Dim col_zoom As Long
        col_zoom = 1
        Do While sht.Cells(1, col_zoom).left < int_cols * cht_wid
            col_zoom = col_zoom + 1
        Loop

        sht.Range("A:A", sht.Cells(1, col_zoom - 1).EntireColumn).Select
        ActiveWindow.Zoom = True
        sht.Range("A1").Select
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

    Dim cht_obj As ChartObject
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        cht_obj.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next cht_obj

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

    'rng_data will contain the block of data with titles included

    Dim rng_data As Range
    Set rng_data = Application.InputBox("Select data with titles", Type:=8)

    Application.ScreenUpdating = False

    Dim iRow As Long, iCol As Long
    iRow = 0

    Dim dHeight As Double, dWidth As Double
    dHeight = 300
    dWidth = 400

    Dim rngColXData As Range, rngColYData As Range
    For Each rngColYData In rng_data.Columns
        iCol = 0

        For Each rngColXData In rng_data.Columns
            If iRow <> iCol Then
                Dim cht As Chart
                Set cht = ActiveSheet.ChartObjects.Add(iCol * dWidth, _
                                                       iRow * dHeight + 100, _
                                                       dWidth, _
                                                       dHeight).Chart

                Dim ser As series
                Dim b_ser As New bUTLChartSeries

                'offset allows for the title to be excluded
                Set b_ser.XValues = Intersect(rngColXData, rngColXData.Offset(1))
                Set b_ser.Values = Intersect(rngColYData, rngColYData.Offset(1))
                Set b_ser.name = rngColYData.Cells(1)
                b_ser.ChartType = xlXYScatter

                Set ser = b_ser.AddSeriesToChart(cht)

                ser.MarkerSize = 3
                ser.MarkerStyle = xlMarkerStyleCircle

                Dim ax As Axis
                Set ax = cht.Axes(xlCategory)
                ax.HasTitle = True
                ax.AxisTitle.Text = rngColXData.Cells(1)
                ax.MajorGridlines.Border.Color = RGB(200, 200, 200)
                ax.MinorGridlines.Border.Color = RGB(220, 220, 220)

                Set ax = cht.Axes(xlValue)
                ax.HasTitle = True
                ax.AxisTitle.Text = rngColYData.Cells(1)
                ax.MajorGridlines.Border.Color = RGB(200, 200, 200)
                ax.MinorGridlines.Border.Color = RGB(220, 220, 220)

                cht.HasTitle = True
                cht.ChartTitle.Text = rngColYData.Cells(1) & " vs. " & rngColXData.Cells(1)
                'cht.ChartTitle.Characters.Font.Size = 8
                cht.Legend.Delete
            End If

            iCol = iCol + 1
        Next

        iRow = iRow + 1
    Next

    Application.ScreenUpdating = True

    rng_data.Cells(1, 1).Activate

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

    Dim cht_obj As ChartObject

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart

        Set cht = cht_obj.Chart

        Dim ser As series
        For Each ser In cht.SeriesCollection

            ser.MarkerSize = 3
            ser.MarkerStyle = xlMarkerStyleCircle

            If ser.ChartType = xlXYScatterLines Then
                ser.Format.Line.Weight = 1.5

            End If

            ser.MarkerForegroundColorIndex = xlColorIndexNone
            ser.MarkerBackgroundColorIndex = xlColorIndexAutomatic

        Next ser


        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionBottom

        Dim ax As Axis
        Set ax = cht.Axes(xlValue)

        ax.MajorGridlines.Border.Color = RGB(242, 242, 242)
        ax.Crosses = xlAxisCrossesMinimum
        
        Set ax = cht.Axes(xlCategory)
        
        ax.HasMajorGridlines = True

        ax.MajorGridlines.Border.Color = RGB(242, 242, 242)

        If cht.HasTitle Then
            cht.ChartTitle.Characters.Font.Size = 12
            cht.ChartTitle.Characters.Font.Bold = True
        End If

        Set ax = cht.Axes(xlCategory)

    Next cht_obj

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

    Dim obj As ChartObject

    For Each obj In Chart_GetObjectsFromObject(Selection)
        obj.Placement = xlFreeFloating
    Next obj

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

    Dim cht_obj As ChartObject


    For Each cht_obj In Selection
        cht_obj.Chart.ChartTitle.Text = cht_obj.Chart.SeriesCollection(1).name
    Next cht_obj


End Sub

