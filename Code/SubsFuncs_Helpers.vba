Attribute VB_Name = "SubsFuncs_Helpers"
''this module contains some common helper code across the add-in

'Helper function which is used to select a range of cells
'This reduces redundant typing redundant typing redundant typing when defining ranges
'+Where is this called? I can't get it to work properly as a UDF on sheet
Function RangeEnd(start As Range, direction As XlDirection, Optional direction2 As XlDirection = -1) As Range
    'check that a second direction was supplied
    If direction2 = -1 Then
        Set RangeEnd = Range(start, start.End(direction))
    Else
        Set RangeEnd = Range(start, start.End(direction).End(direction2))
    End If
End Function

Function RangeEnd_Boundary(start As Range, direction As XlDirection, Optional direction2 As XlDirection = -1) As Range
    '+Where is this called? I can't get it to work properly as a UDF on sheet
    'check that a second direction was supplied
    If direction2 = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(direction)), start.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(start, start.End(direction).End(direction2)), start.CurrentRegion)
    End If
End Function


'from http://stackoverflow.com/a/152325/4288101 & http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
'modified to be case-insensitive and Optional params
Public Sub QuickSort(vArray As Variant, Optional inLow As Variant, Optional inHi As Variant)

    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
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
