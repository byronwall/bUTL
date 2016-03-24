Attribute VB_Name = "SubsFuncs_Helpers"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : SubsFuncs_Helpers
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains some common helper code across the add-in
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : GetInputOrSelection
' Author    : @byronwall
' Date      : 2015 08 11
' Purpose   : Provides a single Function to get the Selection or Input with error handling
'---------------------------------------------------------------------------------------
'
Function GetInputOrSelection(msg As String) As Range
    
    Dim defaultString As String
    
    If TypeOf Selection Is Range Then
        defaultString = Selection.Address
    End If
    
    On Error GoTo ErrorNoSelection
    Set GetInputOrSelection = Application.InputBox(msg, Type:=8, Default:=defaultString)
    
    Exit Function
    
ErrorNoSelection:
    Set GetInputOrSelection = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RangeEnd
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper function to return a block of cells using a starting Range and an End direction
'---------------------------------------------------------------------------------------
'
Function RangeEnd(start As Range, firstDirection As XlDirection, Optional secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd = Range(start, start.End(firstDirection))
    Else
        Set RangeEnd = Range(start, start.End(firstDirection).End(secondDirection))
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : RangeEnd_Boundary
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper function to return a range limited by the starting cell's CurrentRegion
'---------------------------------------------------------------------------------------
'
Function RangeEnd_Boundary(start As Range, firstDirection As XlDirection, Optional secondDirection As XlDirection = -1) As Range

    If secondDirection = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(firstDirection)), start.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(firstDirection).End(secondDirection)), start.CurrentRegion)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : QuickSort
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sorting implementation for arrays
' Source    : http://stackoverflow.com/a/152325/4288101
'             http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
'---------------------------------------------------------------------------------------
'
Public Sub QuickSort(vArray As Variant, Optional incomingLB As Variant, Optional incomingUB As Variant)

    Dim pivot As Variant
    Dim tmpSwap As Variant
    Dim tempLB As Long
    Dim tempUB As Long

    If IsMissing(incomingLB) Then
        incomingLB = LBound(vArray)
    End If

    If IsMissing(incomingUB) Then
        incomingUB = UBound(vArray)
    End If

    tempLB = incomingLB
    tempUB = incomingUB

    pivot = vArray((incomingLB + incomingUB) \ 2)

    While (tempLB <= tempUB)

        While (UCase(vArray(tempLB)) < UCase(pivot) And tempLB < incomingUB)
            tempLB = tempLB + 1
        Wend

        While (UCase(pivot) < UCase(vArray(tempUB)) And tempUB > incomingLB)
            tempUB = tempUB - 1
        Wend

        If (tempLB <= tempUB) Then
            tmpSwap = vArray(tempLB)
            vArray(tempLB) = vArray(tempUB)
            vArray(tempUB) = tmpSwap
            tempLB = tempLB + 1
            tempUB = tempUB - 1
        End If

    Wend

    If (incomingLB < tempUB) Then QuickSort vArray, incomingLB, tempUB
    If (tempLB < incomingUB) Then QuickSort vArray, tempLB, incomingUB

End Sub

