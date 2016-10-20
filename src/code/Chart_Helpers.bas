Attribute VB_Name = "Chart_Helpers"
Option Explicit


Public Function Chart_GetColor(ByVal index As Variant) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetColor
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Returns a list of colors for styling chart series
    '---------------------------------------------------------------------------------------
    '
    Dim colors(1 To 10) As Variant

    colors(1) = RGB(31, 120, 180)
    colors(2) = RGB(227, 26, 28)
    colors(3) = RGB(51, 160, 44)
    colors(4) = RGB(255, 127, 0)
    colors(5) = RGB(106, 61, 154)
    colors(6) = RGB(166, 206, 227)
    colors(7) = RGB(178, 223, 138)
    colors(8) = RGB(251, 154, 153)
    colors(9) = RGB(253, 191, 111)
    colors(10) = RGB(202, 178, 214)

    Chart_GetColor = colors(index)

End Function


Public Function Chart_GetObjectsFromObject(ByVal inputObject As Object) As Variant
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_GetObjectsFromObject
    ' Author    : @byronwall
    ' Date      : 2015 12 30
    ' Purpose   : Helper function which finds a valid ChartObject based on what is actually selected
    '             Returns a Collection (possibly empty) and should be handled with a For Each
    '---------------------------------------------------------------------------------------
    '
    Dim chartObjectCollection As New Collection

    'NOTE that this function does not work well with Axis objects.  Excel does not return the correct Parent for them.
    
    Dim targetObject As Variant
    Dim inputObjectType As String
    inputObjectType = TypeName(inputObject)

    Select Case inputObjectType
    
        Case "DrawingObjects"
            'this means that multiple charts are selected
            For Each targetObject In inputObject
                If TypeName(targetObject) = "ChartObject" Then
                    'add it to the set
                    chartObjectCollection.Add targetObject
                End If
            Next targetObject
            
        Case "Worksheet"
            For Each targetObject In inputObject.ChartObjects
                chartObjectCollection.Add targetObject
            Next targetObject
            
        Case "Chart"
            chartObjectCollection.Add inputObject.Parent
            
        Case "ChartArea", "PlotArea", "Legend", "ChartTitle"
            'parent is the chart, parent of that is the chart targetObject
            chartObjectCollection.Add inputObject.Parent.Parent
            
        Case "series"
            'need to go up three levels
            chartObjectCollection.Add inputObject.Parent.Parent.Parent
            
        Case "Axis", "Gridlines", "AxisTitle"
            'these are the oddly unsupported objects
            MsgBox "Axis/gridline selection not supported.  This is an Excel bug.  Select another element on the chart(s)."
    
        Case Else
            MsgBox "Select a part of the chart(s), except an axis."
    
    End Select

    Set Chart_GetObjectsFromObject = chartObjectCollection
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

        Dim chartObjectIndex As Long
        For chartObjectIndex = ActiveSheet.ChartObjects.count To 1 Step -1

            ActiveSheet.ChartObjects(chartObjectIndex).Delete

        Next chartObjectIndex

        Application.ScreenUpdating = True

    End If
End Sub

