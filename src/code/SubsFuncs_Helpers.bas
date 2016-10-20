Attribute VB_Name = "SubsFuncs_Helpers"
Option Explicit


Public Function GetInputOrSelection(ByVal userPrompt As String) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : GetInputOrSelection
    ' Author    : @byronwall
    ' Date      : 2015 08 11
    ' Purpose   : Provides a single Function to get the Selection or Input with error handling
    '---------------------------------------------------------------------------------------
    '
    Dim defaultString As String
    
    If TypeOf Selection Is Range Then
        defaultString = Selection.Address
    End If
    
    On Error GoTo ErrorNoSelection
    Set GetInputOrSelection = Application.InputBox(userPrompt, Type:=8, Default:=defaultString)
    
    Exit Function
    
ErrorNoSelection:
    Set GetInputOrSelection = Nothing
    
End Function



Public Function RangeEnd(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : RangeEnd
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Helper function to return a block of cells using a starting Range and an End firstDirection
    '---------------------------------------------------------------------------------------
    '
    If secondDirection = -1 Then
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection))
    Else
        Set RangeEnd = Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection))
    End If
End Function


Public Function RangeEnd_Boundary(ByVal rangeBegin As Range, ByVal firstDirection As XlDirection, Optional ByVal secondDirection As XlDirection = -1) As Range
    '---------------------------------------------------------------------------------------
    ' Procedure : RangeEnd_Boundary
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Helper function to return a range limited by the starting cell's CurrentRegion
    '---------------------------------------------------------------------------------------
    '
    If secondDirection = -1 Then
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection)), rangeBegin.CurrentRegion)
    Else
        Set RangeEnd_Boundary = Intersect(Range(rangeBegin, rangeBegin.End(firstDirection).End(secondDirection)), rangeBegin.CurrentRegion)
    End If
End Function


Public Sub QuickSort(ByVal arrayToSort As Variant, Optional ByVal lowBound As Variant, Optional ByVal highBound As Variant)
    '---------------------------------------------------------------------------------------
    ' Procedure : QuickSort
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Sorting implementation for arrays
    ' Source    : http://stackoverflow.com/a/152325/4288101
    '             http://en.allexperts.com/q/Visual-Basic-1048/string-manipulation.htm
    '---------------------------------------------------------------------------------------
    '
    Dim sortingVariant As Variant
    Dim swapHolder As Variant
    Dim temporaryLowBound As Long
    Dim temporaryHighBound As Long

    If IsMissing(lowBound) Then lowBound = LBound(arrayToSort)
    If IsMissing(highBound) Then highBound = UBound(arrayToSort)

    temporaryLowBound = lowBound
    temporaryHighBound = highBound

    sortingVariant = arrayToSort((lowBound + highBound) \ 2)

    While (temporaryLowBound <= temporaryHighBound)

        While (UCase(arrayToSort(temporaryLowBound)) < UCase(sortingVariant) And temporaryLowBound < highBound)
            temporaryLowBound = temporaryLowBound + 1
        Wend

        While (UCase(sortingVariant) < UCase(arrayToSort(temporaryHighBound)) And temporaryHighBound > lowBound)
            temporaryHighBound = temporaryHighBound - 1
        Wend

        If (temporaryLowBound <= temporaryHighBound) Then
            swapHolder = arrayToSort(temporaryLowBound)
            arrayToSort(temporaryLowBound) = arrayToSort(temporaryHighBound)
            arrayToSort(temporaryHighBound) = swapHolder
            temporaryLowBound = temporaryLowBound + 1
            temporaryHighBound = temporaryHighBound - 1
        End If

    Wend

    If (lowBound < temporaryHighBound) Then QuickSort arrayToSort, lowBound, temporaryHighBound
    If (temporaryLowBound < highBound) Then QuickSort arrayToSort, temporaryLowBound, highBound

End Sub

