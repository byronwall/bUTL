Attribute VB_Name = "Chart_Format"
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

Sub Chart_ApplyFormattingToSelected()

    Dim obj As ChartObject

    For Each obj In Chart_GetObjectsFromObject(Selection)
        
        Dim ser As series
        
        For Each ser In obj.Chart.SeriesCollection
            ser.MarkerSize = 5
        Next ser
    Next obj
    
End Sub

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

Sub Chart_AxisTitleIsSeriesTitle()
    'this will make the axis titles equal to the names of the series on them (last one wins)
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
            
        Next ser
    Next cht_obj
End Sub

Sub Chart_CreateDataLabels()

    Dim cht As Chart
    
    Set cht = Selection.Chart
    
    Dim ser As series
    
    For Each ser In cht.SeriesCollection
        
        Dim p As Point
        Set p = ser.Points(2)
        
        p.HasDataLabel = False
        p.DataLabel.Position = xlLabelPositionRight
        p.DataLabel.ShowSeriesName = True
        p.DataLabel.ShowValue = False
        p.DataLabel.ShowCategoryName = False
        p.DataLabel.ShowLegendKey = True
            
    Next ser

End Sub

''this is the business end of creating a grid of charts
Sub Chart_GridOfCharts( _
    Optional int_cols As Integer = 3, _
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
    
    Dim count As Integer
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
        cht_obj.height = cht_height
        
        count = count + 1
        
    Next cht_obj
    
    'loop through columsn to find how far to zoom
    If bool_zoom Then
        Dim col_zoom As Integer
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

Sub ChartApplyToAll()

    Dim cht_obj As ChartObject
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        cht_obj.Chart.SeriesCollection(1).ChartType = xlXYScatter
    Next cht_obj

End Sub

Sub ChartCreateXYGrid()
    
    'delete all charts?
    If MsgBox("Delete all charts?", vbYesNo) = vbYes Then
        Application.ScreenUpdating = False
        Dim cht_obj As ChartObject
        For Each cht_obj In ActiveSheet.ChartObjects
            cht_obj.Delete
        Next cht_obj
        
        Application.ScreenUpdating = True
        
    End If

    'rng_data will contain the block of data with titles included
    Dim rng_data As Range
    Set rng_data = Application.InputBox("Select data with titles", Type:=8)
    
    Application.ScreenUpdating = False
    
    'iterate through each column
    'iterate through each column again
    
    'create a chart
    'set the x data to one column, y data to the other
    
    Dim int_row As Integer, int_col As Integer
    int_row = 0

    
    Dim dbl_height As Double, dbl_width As Double
    dbl_height = 300
    dbl_width = 400
    
    Dim rng_colXData As Range, rng_colYData As Range
    For Each rng_colYData In rng_data.Columns
        int_col = 0
        
        For Each rng_colXData In rng_data.Columns
        
            If int_row = int_col Then
            
            Else
            
                Dim cht As Chart
                Set cht = ActiveSheet.ChartObjects.Add(int_col * dbl_width, int_row * dbl_height + 100, dbl_width, dbl_height).Chart
                
                Dim ser As series
                Dim b_ser As New bUTLChartSeries
                
                'offset allows for the title to be excluded
                Set b_ser.XValues = Intersect(rng_colXData, rng_colXData.Offset(1))
                Set b_ser.Values = Intersect(rng_colYData, rng_colYData.Offset(1))
                Set b_ser.name = rng_colYData.Cells(1)
                b_ser.ChartType = xlXYScatter
                
                Set ser = b_ser.AddSeriesToChart(cht)
                
                ser.MarkerSize = 3
                ser.MarkerStyle = xlMarkerStyleCircle
                
                Dim ax As Axis
                Set ax = cht.Axes(xlCategory)
                    ax.HasTitle = True
                    ax.AxisTitle.Text = rng_colXData.Cells(1)
                    ax.MajorGridlines.Border.Color = RGB(200, 200, 200)
                    ax.MinorGridlines.Border.Color = RGB(220, 220, 220)
                
                Set ax = cht.Axes(xlValue)
                    ax.HasTitle = True
                    ax.AxisTitle.Text = rng_colYData.Cells(1)
                    ax.MajorGridlines.Border.Color = RGB(200, 200, 200)
                    ax.MinorGridlines.Border.Color = RGB(220, 220, 220)
                
                cht.HasTitle = True
                cht.ChartTitle.Text = rng_colYData.Cells(1) & " vs. " & rng_colXData.Cells(1)
                'cht.ChartTitle.Characters.Font.Size = 8
                cht.Legend.Delete
            End If
            
            int_col = int_col + 1
            
            
        Next
            
            
        int_row = int_row + 1
    Next
    
    Application.ScreenUpdating = True
    
    rng_data.Cells(1, 1).Activate

End Sub

Sub ChartDefaultFormat()

    Dim cht_obj As ChartObject
    
    For Each cht_obj In ActiveSheet.ChartObjects
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
        
        ax.MajorGridlines.Border.Color = RGB(200, 200, 200)
        ax.MinorGridlines.Border.Color = RGB(230, 230, 230)
        ax.Crosses = xlAxisCrossesMinimum
        
        If cht.HasTitle Then
            cht.ChartTitle.Characters.Font.Size = 12
            cht.ChartTitle.Characters.Font.Bold = True
        End If
        
        Set ax = cht.Axes(xlCategory)
               
    Next cht_obj

End Sub

Sub ChartTitleEqualsSeriesSelection()

    Dim cht_obj As ChartObject
    

        For Each cht_obj In Selection
            cht_obj.Chart.ChartTitle.Text = cht_obj.Chart.SeriesCollection(1).name
        Next cht_obj


End Sub

Sub ChartPropMove()

    Dim obj As ChartObject

    For Each obj In Chart_GetObjectsFromObject(Selection)
        obj.Placement = xlFreeFloating
    Next obj
    
End Sub




