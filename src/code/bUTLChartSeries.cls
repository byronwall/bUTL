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

Private str_name As String

'---------------------------------------------------------------------------------------
' Procedure : AddSeriesToChart
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Adds the represented series to a chart
'---------------------------------------------------------------------------------------
'
Function AddSeriesToChart(cht As Chart) As series

    Dim ser As series
    Set ser = cht.SeriesCollection.newSeries
    
    ser.Formula = Me.SeriesFormula
    
    If Me.ChartType <> 0 Then
        ser.ChartType = Me.ChartType
    End If
    
    
    Set AddSeriesToChart = ser

End Function

'---------------------------------------------------------------------------------------
' Procedure : FullAddress
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Gets the full address for a range
'---------------------------------------------------------------------------------------
'
Private Function FullAddress(rng As Range) As Variant

    If rng Is Nothing Then
        FullAddress = ""
    Else
        FullAddress = "'" & rng.Parent.name & "'!" & rng.Address(External:=False)
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
    If str_name <> "" Then
        SeriesFormula = "=SERIES(" & str_name & "," & _
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
Sub UpdateFromChartSeries(ser As series)
'this will work for the simple case where all items are references

    Set series = ser

    Dim form As Variant

    '=SERIES("Y",Sheet1!$C$8:$C$13,Sheet1!$D$8:$D$13,1)

    'pull in teh formula
    form = ser.Formula

    'uppercase to remove match errors
    form = UCase(form)

    'remove the front of the formula
    form = Replace(form, "=SERIES(", "")

    'find the first comma
    Dim comma
    comma = InStr(form, ",")

    If comma > 1 Then

        'need to catch an error here if a text name is used instead of a valid range
        On Error Resume Next
        Set Me.name = Range(left(form, comma - 1))

        If Err <> 0 Then
            str_name = left(form, comma - 1)
        End If

        On Error GoTo 0
    End If

    'pull out the title from that
    form = Mid(form, comma + 1)

    comma = InStr(form, ",")

    If comma > 1 Then
        Set Me.XValues = Range(left(form, comma - 1))
    End If
    form = Mid(form, comma + 1)

    comma = InStr(form, ",")
    Set Me.Values = Range(left(form, comma - 1))
    form = Mid(form, comma + 1)

    comma = InStr(form, ")")
    Me.SeriesNumber = left(form, comma - 1)

    Me.ChartType = ser.ChartType
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

