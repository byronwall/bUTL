Attribute VB_Name = "UDFs"
Option Explicit


Public Function RandLetters(num As Long) As String
    '---------------------------------------------------------------------------------------
    ' Procedure : RandLetters
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : UDF that generates a sequence of random letters
    '---------------------------------------------------------------------------------------
    '
    Dim i As Long
    
    Dim letters() As String
    ReDim letters(1 To num)
    
    For i = 1 To num
        letters(i) = Chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function

