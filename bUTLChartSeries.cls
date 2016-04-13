VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bUTLChartSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : bUTLChartSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Helper class to read series from charts and manipulate easier
'---------------------------------------------------------------------------------------

Public Values As Range
Public XValues As Range
Public name As Range
Public SeriesNumber As Long
Public ChartType As XlChartType
Public series As series

Private StringName As String

'---------------------------------------------------------------------------------------
' Procedure : AddSeriesToChart
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds the represented series to a chart
'---------------------------------------------------------------------------------------
'
Function AddSeriesToChart(targetChart As Chart) As series

    Dim mySeries As series
    Set mySeries = targetChart.SeriesCollection.newSeries
    
    mySeries.Formula = Me.SeriesFormula
    
    If Me.ChartType <> 0 Then
        mySeries.ChartType = Me.ChartType
    End If
    
    
    Set AddSeriesToChart = mySeries

End Function

'---------------------------------------------------------------------------------------
' Procedure : FullAddress
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Gets the full address for a range
'---------------------------------------------------------------------------------------
'
Private Function FullAddress(targetRange As Range) As Variant

    If targetRange Is Nothing Then
        FullAddress = ""
    Else
        FullAddress = "'" & targetRange.Parent.name & "'!" & targetRange.Address(External:=False)
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : SeriesFormula
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Returns a SERIES formula for the represented series
'---------------------------------------------------------------------------------------
'
Public Property Get SeriesFormula() As String

    '2015 11 09 add a trap here to allow for a string only name
    If StringName <> "" Then
        SeriesFormula = "=SERIES(" & StringName & "," & _
                        FullAddress(Me.XValues) & "," & FullAddress(Me.Values) & "," & _
                        Me.SeriesNumber & ")"
    Else

        SeriesFormula = "=SERIES(" & FullAddress(Me.name) & "," & _
                        FullAddress(Me.XValues) & "," & FullAddress(Me.Values) & "," & _
                        Me.SeriesNumber & ")"
    End If

End Property

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Ensures the SERIES starts out first
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Me.SeriesNumber = 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateFromChartSeries
' Author    : @byronwall
' Date      : 2015 11 09
' Purpose   : Reads the series info from a Series and stores it in the class
'---------------------------------------------------------------------------------------
'
Sub UpdateFromChartSeries(mySeries As series)
'this will work for the simple case where all items are references

    Set series = mySeries

    Dim myForm As Variant

    '=SERIES("Y",Sheet1!$C$8:$C$13,Sheet1!$D$8:$D$13,1)

    'pull in teh formula
    myForm = mySeries.Formula

    'uppercase to remove match errors
    myForm = UCase(myForm)

    'remove the front of the formula
    myForm = Replace(myForm, "=SERIES(", "")

    'find the first delimiterPosition
    Dim delimiterPosition As Long
    delimiterPosition = InStr(myForm, ",")

    If delimiterPosition > 1 Then

        'need to catch an error here if a text name is used instead of a valid range
        On Error Resume Next
        Set Me.name = Range(left(myForm, delimiterPosition - 1))

        If Err <> 0 Then
            StringName = left(myForm, delimiterPosition - 1)
        End If

        On Error GoTo 0
    End If

    'pull out the title from that
    myForm = Mid(myForm, delimiterPosition + 1)

    delimiterPosition = InStr(myForm, ",")

    If delimiterPosition > 1 Then
        Set Me.XValues = Range(left(myForm, delimiterPosition - 1))
    End If
    myForm = Mid(myForm, delimiterPosition + 1)

    delimiterPosition = InStr(myForm, ",")
    Set Me.Values = Range(left(myForm, delimiterPosition - 1))
    myForm = Mid(myForm, delimiterPosition + 1)

    delimiterPosition = InStr(myForm, ")")
    Me.SeriesNumber = left(myForm, delimiterPosition - 1)

    Me.ChartType = mySeries.ChartType
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateSeriesWithNewValues
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Forces the stored Series to take on any changed values in the class
'---------------------------------------------------------------------------------------
'
Sub UpdateSeriesWithNewValues()

    Me.series.Formula = Me.SeriesFormula

End Sub

