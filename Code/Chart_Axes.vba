Attribute VB_Name = "Chart_Axes"
'this contains code that related to charting axes

Sub Chart_Axis_AutoY()
   
    Dim cht_obj As ChartObject
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart
        
        Dim ax As Axis
        
        Set cht = cht_obj.Chart
        
        Set ax = cht.Axes(xlValue)
        ax.MaximumScaleIsAuto = True
        ax.MinimumScaleIsAuto = True
        ax.MajorUnitIsAuto = True
        ax.MinorUnitIsAuto = True
    
    Next cht_obj

End Sub

Sub Chart_Axis_AutoX()
   
    Dim cht_obj As ChartObject
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart
        
        Dim ax As Axis
        
        Set cht = cht_obj.Chart
        
        Set ax = cht.Axes(xlCategory)
        ax.MaximumScaleIsAuto = True
        ax.MinimumScaleIsAuto = True
        ax.MajorUnitIsAuto = True
        ax.MinorUnitIsAuto = True
    
    Next cht_obj

End Sub

Sub Chart_YAxisRangeWithAvgAndStdev()
    'this sub will set the y-axis to a number of stdevs past the average of the first series
    Dim dbl_std As Double
    
    dbl_std = CDbl(InputBox("How many standard deviations to include?"))
    
    Dim cht_obj As ChartObject
    
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        
        Dim ser As series
        Set ser = cht_obj.Chart.SeriesCollection(1)
            
        Dim avg_val As Double
        Dim std_val As Double
        
        avg_val = WorksheetFunction.Average(ser.Values)
        std_val = WorksheetFunction.StDev(ser.Values)
                
        cht_obj.Chart.Axes(xlValue).MinimumScale = avg_val - std_val * dbl_std
        cht_obj.Chart.Axes(xlValue).MaximumScale = avg_val + std_val * dbl_std
    
    Next

End Sub

Sub Chart_FitAxisToMaxAndMin(xlCat As XlAxisType)
    'this will loop through all charts in the selection and process the axes on them
    
    Dim first As Boolean
    first = True
    
    Dim cht_obj As ChartObject
    
    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        Dim cht As Chart
        Set cht = cht_obj.Chart
        
        Dim ser As series
        For Each ser In cht.SeriesCollection
            
            Dim min_val As Double, max_val As Double
            
            If xlCat = xlCategory Then
            
                min_val = Application.Min(ser.XValues)
                max_val = Application.Max(ser.XValues)
            
            ElseIf xlCat = xlValue Then
            
                min_val = Application.Min(ser.Values)
                max_val = Application.Max(ser.Values)
            
            End If
            
                    
            Dim ax As Axis
            Set ax = cht.Axes(xlCat)
            
            Dim bool_max As Boolean, bool_min As Boolean
            bool_max = max_val > ax.MaximumScale
            bool_min = min_val < ax.MinimumScale
            
            If first Or bool_min Then
                ax.MinimumScale = min_val
            End If
            If first Or bool_max Then
                ax.MaximumScale = max_val
            End If
            
            first = False
        Next ser
    Next cht_obj

End Sub
