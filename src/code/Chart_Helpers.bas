Attribute VB_Name = "Chart_Helpers"
Option Explicit


Public Function Chart_GetColor(index As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetColor
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Returns a list of colors for styling chart series
    '---------------------------------------------------------------------------------------
    '
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


Public Function Chart_GetObjectsFromObject(obj_in As Object) As Variant
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetObjectsFromObject
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper function which finds a valid ChartObject based on what is actually selected
    '             Returns a Collection (possibly empty) and should be handled with a For Each
    '---------------------------------------------------------------------------------------
    '
    Dim chtObjCollection As New Collection

    'NOTE that this function does not work well with Axis objects.  Excel does not return the correct Parent for them.
    
    Dim obj As Variant


    If TypeOf obj_in Is DrawingObjects Then
        'this means that multiple charts are selected
        
        For Each obj In obj_in
            If TypeName(obj) = "ChartObject" Then
                'add it to the set
                chtObjCollection.Add obj
            End If
        Next obj
        
    ElseIf TypeOf obj_in Is Worksheet Then
        For Each obj In obj_in.ChartObjects
            chtObjCollection.Add obj
        Next obj

    ElseIf TypeOf obj_in Is Chart Then
        chtObjCollection.Add obj_in.Parent

    ElseIf TypeOf obj_in Is ChartArea _
           Or TypeOf obj_in Is PlotArea _
           Or TypeOf obj_in Is Legend _
           Or TypeOf obj_in Is ChartTitle Then
        'parent is the chart, parent of that is the chart obj
        chtObjCollection.Add obj_in.Parent.Parent

    ElseIf TypeOf obj_in Is series Then
        'need to go up three levels
        chtObjCollection.Add obj_in.Parent.Parent.Parent

    ElseIf TypeOf obj_in Is Axis _
           Or TypeOf obj_in Is Gridlines _
           Or TypeOf obj_in Is AxisTitle Then
        'these are the oddly unsupported objects
        MsgBox "Axis/gridline selection not supported.  This is an Excel bug.  Select another element on the chart(s)."

    Else
        MsgBox "Select a part of the chart(s), except an axis."

    End If

    Set Chart_GetObjectsFromObject = chtObjCollection
End Function


Public Sub DeleteAllCharts()
    '---------------------------------------------------------------------------------------
    ' Procedure : DeleteAllCharts
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Helper Sub to delete all charts on ActiveSheet
    '---------------------------------------------------------------------------------------
    '
    If MsgBox("Delete all charts?", vbYesNo) = vbYes Then
        Application.ScreenUpdating = False

        Dim chtObjIndex As Long
        For chtObjIndex = ActiveSheet.ChartObjects.count To 1 Step -1

            ActiveSheet.ChartObjects(chtObjIndex).Delete

        Next chtObjIndex

        Application.ScreenUpdating = True

    End If
End Sub

