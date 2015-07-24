Attribute VB_Name = "Chart_Helpers"
'this contains VBA that is used across charting modules to help things go

'this returns the bUTL colors
Function Chart_GetColor(index As Variant) As Long

    Dim colors(1 To 10) As Variant
    
    colors(6) = RGB(166, 206, 227)
    colors(1) = RGB(31, 120, 180)
    colors(7) = RGB(178, 223, 138)
    colors(3) = RGB(51, 160, 44)
    colors(8) = RGB(251, 154, 153)
    colors(2) = RGB(227, 26, 28)
    colors(9) = RGB(253, 191, 111)
    colors(4) = RGB(255, 127, 0)
    colors(10) = RGB(202, 178, 214)
    colors(5) = RGB(106, 61, 154)
    
    Chart_GetColor = colors(index)


End Function
Function Chart_GetObjectsFromObject(obj_in As Object) As Variant
    
    Dim str_type As String
    str_type = TypeName(obj_in)
    
    Dim coll As New Collection
    
    Dim obj As Variant
    
    If str_type = "DrawingObjects" Then
        'this means that multiple charts are selected
        For Each obj In obj_in
            If TypeName(obj) = "ChartObject" Then
                'add it to the set
                coll.Add obj
            End If
        Next obj
    
    ElseIf str_type = "Chart" Then
        coll.Add obj_in.Parent
    
    ElseIf str_type = "ChartArea" Or str_type = "PlotArea" Then
        'parent is the chart, parent of that is the chart obj
        coll.Add obj_in.Parent.Parent
    
    ElseIf str_type = "Series" Then
        'need to go up three levels
        coll.Add obj_in.Parent.Parent.Parent
        
    Else
        MsgBox "Select an object that is supported."
    End If
    
    Set Chart_GetObjectsFromObject = coll

End Function



