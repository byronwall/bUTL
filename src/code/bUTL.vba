VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bUTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : bUTL
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Holder class that allows some of the Ribbon features to have a state
'---------------------------------------------------------------------------------------

Option Explicit

Private txt_values As New Dictionary

'---------------------------------------------------------------------------------------
' Procedure : GetTextValue
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Gets a value from the stored dictionary
'---------------------------------------------------------------------------------------
'
Public Function GetTextValue(id As Variant) As Variant
    If txt_values.Exists(id) Then
        GetTextValue = txt_values(id)
    Else
        GetTextValue = Null
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetTextValue
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Sets a dictionary value, used with the text boxes
'---------------------------------------------------------------------------------------
'
Public Sub SetTextValue(id As Variant, Text As Variant)
    txt_values(id) = Text
End Sub

