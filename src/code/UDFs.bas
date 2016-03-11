Attribute VB_Name = "UDFs"
'---------------------------------------------------------------------------------------
' Module    : UDFs
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains all code that is intended to be used as a UDF
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : RandLetters
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : UDF that generates a sequence of random letters
'---------------------------------------------------------------------------------------
'
Public Function RandLetters(numberOfLetters As Long) As String

    Dim i As Long
    
    Dim letters() As String
    ReDim letters(1 To numberOfLetters)
    
    For i = 1 To numberOfLetters
        letters(i) = Chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function

