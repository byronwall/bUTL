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

Private pName As String


Public Function AddSeriesToChart(ByVal targetChart As Chart) As series
    '---------------------------------------------------------------------------------------
    ' Procedure : AddSeriesToChart
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Adds the represented series to a chart
    '---------------------------------------------------------------------------------------
    '
    Dim targetSeries As series
    Set targetSeries = targetChart.SeriesCollection.NewSeries
    
    targetSeries.Formula = Me.SeriesFormula
    
    If Me.ChartType <> 0 Then targetSeries.ChartType = Me.ChartType
    
    Set AddSeriesToChart = targetSeries

End Function


Private Function FullAddress(targetRange As Range) As Variant
    '---------------------------------------------------------------------------------------
    ' Procedure : FullAddress
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Gets the full address for a range
    '---------------------------------------------------------------------------------------
    '
    Const START_SHEET_LINKER As String = "'"
    Const END_SHEET_LINKER As String = "'!"
    If targetRange Is Nothing Then
        FullAddress = vbNullString
    Else
        FullAddress = START_SHEET_LINKER & targetRange.Parent.name & END_SHEET_LINKER & targetRange.Address(External:=False)
    End If

End Function


Public Property Get SeriesFormula() As String
    '---------------------------------------------------------------------------------------
    ' Procedure : SeriesFormula
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Returns a SERIES formula for the represented series
    '---------------------------------------------------------------------------------------
    '
    '2015 11 09 add a trap here to allow for a string only name
    Const COMMA As String = ","
    Const START_FORMULA As String = "=SERIES("
    Const END_FORMULA As String = ")"
    
    If pName <> vbNullString Then
        SeriesFormula = START_FORMULA & pName & COMMA & _
                        FullAddress(Me.XValues) & COMMA & FullAddress(Me.Values) & COMMA & _
                        Me.SeriesNumber & END_FORMULA
    Else

        SeriesFormula = START_FORMULA & FullAddress(Me.name) & COMMA & _
                        FullAddress(Me.XValues) & COMMA & FullAddress(Me.Values) & COMMA & _
                        Me.SeriesNumber & END_FORMULA
    End If

End Property


Private Sub Class_Initialize()
    '---------------------------------------------------------------------------------------
    ' Procedure : Class_Initialize
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Ensures the SERIES starts out first
    '---------------------------------------------------------------------------------------
    '
    Me.SeriesNumber = 1
End Sub


Public Sub UpdateFromChartSeries(targetSeries As series)
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateFromChartSeries
    ' Author    : @byronwall
    ' Date      : 2015 11 09
    ' Purpose   : Reads the series info from a Series and stores it in the class
    '---------------------------------------------------------------------------------------
    '
    
    'this will work for the simple case where all items are references
    Const FIND_STRING As String = "SERIES("
    Const COMMA As String = ","
    Const CLOSE_BRACKET As String = ")"
    
    Set series = targetSeries

    Dim targetForm As Variant

    '=SERIES("Y",Sheet1!$C$8:$C$13,Sheet1!$D$8:$D$13,1)

    'pull in teh formula
    targetForm = targetSeries.Formula

    'uppercase to remove match errors
    targetForm = UCase(targetForm)

    'remove the front of the formula
    targetForm = Replace(targetForm, FIND_STRING, vbNullString)
    
    'find the first foundPosition
    Dim foundPosition As Long
    foundPosition = InStr(targetForm, COMMA)

    If foundPosition > 1 Then
        'need to catch an error here if a text name is used instead of a valid range
        On Error Resume Next
        Set Me.name = Range(left(targetForm, foundPosition - 1))
        If Err <> 0 Then pName = left(targetForm, foundPosition - 1)
        On Error GoTo 0
    End If

    'pull out the title from that
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, COMMA)

    If foundPosition > 1 Then Set Me.XValues = Range(left(targetForm, foundPosition - 1))
 
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, COMMA)
    Set Me.Values = Range(left(targetForm, foundPosition - 1))
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, CLOSE_BRACKET)
    Me.SeriesNumber = left(targetForm, foundPosition - 1)

    Me.ChartType = targetSeries.ChartType
End Sub


Public Sub UpdateSeriesWithNewValues()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateSeriesWithNewValues
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Forces the stored Series to take on any changed values in the class
    '---------------------------------------------------------------------------------------
    '
    Me.series.Formula = Me.SeriesFormula
End Sub



