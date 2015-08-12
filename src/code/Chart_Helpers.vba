Attribute VB_Name = "Chart_Helpers"
'---------------------------------------------------------------------------------------
' Module    : Chart_Helpers
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains code that helps other chart related features
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : Chart_GetColor
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Returns a list of colors for styling chart series
'---------------------------------------------------------------------------------------
'
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

'---------------------------------------------------------------------------------------
' Procedure : Chart_GetObjectsFromObject
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper function which finds a valid ChartObject based on what is actually selected
'             Returns a Collection (possibly empty) and should be handled with a For Each
'---------------------------------------------------------------------------------------
'
Function Chart_GetObjectsFromObject(obj_in As Object) As Variant

    Dim str_type As String
    'TODO: these should be upgrade to TypeOf instead of strings
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

'---------------------------------------------------------------------------------------
' Procedure : DeleteAllCharts
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Helper Sub to delete all charts on ActiveSheet
'---------------------------------------------------------------------------------------
'
Sub DeleteAllCharts()

    If MsgBox("Delete all charts?", vbYesNo) = vbYes Then
        Application.ScreenUpdating = False

        Dim iCounter As Integer
        For iCounter = ActiveSheet.ChartObjects.count To 1 Step -1

            ActiveSheet.ChartObjects(iCounter).Delete

        Next iCounter

        Application.ScreenUpdating = True

    End If
End Sub

