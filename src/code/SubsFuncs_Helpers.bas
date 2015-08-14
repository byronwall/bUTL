Attribute VB_Name = "SubsFuncs_Helpers"
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
    
    Dim strDefault As String
    
    If TypeOf Selection Is Range Then
        strDefault = Selection.Address
    End If
    
    On Error GoTo ErrorNoSelection
    Set GetInputOrSelection = Application.InputBox(msg, Type:=8, Default:=strDefault)
    
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
Function RangeEnd(start As Range, direction As XlDirection, Optional direction2 As XlDirection = -1) As Range

    If direction2 = -1 Then
        Set RangeEnd = Range(start, start.End(direction))
    Else
        Set RangeEnd = Range(start, start.End(direction).End(direction2))
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : RangeEnd_Boundary
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper function to return a range limited by the starting cell's CurrentRegion
'---------------------------------------------------------------------------------------
'
Function RangeEnd_Boundary(start As Range, direction As XlDirection, Optional direction2 As XlDirection = -1) As Range

    If direction2 = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(direction)), start.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(direction).End(direction2)), start.CurrentRegion)
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
Public Sub QuickSort(vArray As Variant, Optional inLow As Variant, Optional inHi As Variant)

    Dim pivot As Variant
    Dim tmpSwap As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long

    If IsMissing(inLow) Then
        inLow = LBound(vArray)
    End If

    If IsMissing(inHi) Then
        inHi = UBound(vArray)
    End If

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)

        While (UCase(vArray(tmpLow)) < UCase(pivot) And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (UCase(pivot) < UCase(vArray(tmpHi)) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If

    Wend

    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

