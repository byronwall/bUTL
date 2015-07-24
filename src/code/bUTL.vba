VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bUTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private txt_values As New Dictionary

Public Sub SetTextValue(id As Variant, Text As Variant)
    txt_values(id) = Text
End Sub

Public Function GetTextValue(id As Variant) As Variant
    If txt_values.Exists(id) Then
        GetTextValue = txt_values(id)
    Else
        GetTextValue = Null
    End If
End Function




