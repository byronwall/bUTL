Attribute VB_Name = "Chart_Unknown"
'''''
' this contains a lot of "extra" code that is not reference by any buttons
' some this is useful and the rest is junk

Sub ChartSeriesColoring()

    Dim cht As Chart
    Dim cht_obj As ChartObject
    
    'assumes a chart object is selected
    Set cht_obj = Selection
    Set cht = cht_obj.Chart
    
    Dim int_colors As Integer
    int_colors = cht.SeriesCollection.count
    
    Dim int_maxColor As Integer
    Dim int_minColor As Integer
    
    int_maxColor = 255
    int_minColor = 100
    
    Dim int_index As Integer
    int_index = 0
    
    Dim int_indexDelta As Integer
    int_indexDelta = 0
    
    Dim ser As series
    For Each ser In cht.SeriesCollection
        
        Dim int_color As Integer
        
        If int_index / (int_colors / 3) < 1 Then
        ElseIf int_index / (int_colors / 3) < 2 Then
            int_indexDelta = int_colors / 3
        Else
            int_indexDelta = int_colors * 2 / 3
        End If
        
        int_color = int_maxColor - (int_maxColor - int_minColor) / int_colors * (int_index - int_indexDelta) * 3
                
        If int_index / (int_colors / 3) < 1 Then
            ser.MarkerBackgroundColor = RGB(int_color, 0, 0)
        ElseIf int_index / (int_colors / 3) < 2 Then
            ser.MarkerBackgroundColor = RGB(0, int_color, 0)
        Else
            ser.MarkerBackgroundColor = RGB(0, 0, int_color)
        End If
        
        int_index = int_index + 1
        
    Next ser
    

End Sub
