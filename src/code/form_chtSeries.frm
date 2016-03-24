VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_chtSeries 
   Caption         =   "Correct Series"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   OleObjectBlob   =   "form_chtSeries.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_chtSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'---------------------------------------------------------------------------------------
' Module    : form_chtSeries
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Code is under development to better change multiple series at once
'---------------------------------------------------------------------------------------


Dim ser_coll As New Dictionary
Dim dirty As Boolean

Private Sub btn_setXRange_Click()

'get the selected series

    Dim i As Integer
    For i = 0 To list_series.ListCount - 1

        If list_series.Selected(i) Then

            Dim b_ser As bUTLChartSeries
            Set b_ser = ser_coll(i & list_series.List(i, 0))

            Set b_ser.XValues = Range(txt_xrange)

            b_ser.series.Formula = b_ser.SeriesFormula

        End If

    Next i

    UpdateSeries

End Sub

Private Sub btn_xrangedown_Click()

    txt_xrange = RangeEnd(Range(txt_xrange), xlDown).Address(, , , True)

End Sub

Private Sub btn_ydown_Click()
    txt_yrange = RangeEnd(Range(txt_yrange), xlDown).Address(, , , True)
End Sub

Private Sub btn_yrange_Click()
    Dim i As Integer
    For i = 0 To list_series.ListCount - 1

        If list_series.Selected(i) Then

            Dim b_ser As bUTLChartSeries
            Set b_ser = ser_coll(i & list_series.List(i, 0))

            Set b_ser.Values = Range(txt_yrange)

            b_ser.series.Formula = b_ser.SeriesFormula

        End If

    Next i

    UpdateSeries
End Sub

Private Sub txt_xrange_Enter()
    Hide

    Dim rng_data As Range
    Set rng_data = Application.InputBox("Select start", Type:=8)

    Show

    txt_xrange = rng_data.Address(External:=True)
End Sub

Private Sub txt_yrange_Enter()
    Hide

    Dim rng_data As Range
    Set rng_data = Application.InputBox("Select start", Type:=8)

    Show

    txt_yrange = rng_data.Address(External:=True)
End Sub

Private Sub UpdateSeries()

'clean up the mess
    ser_coll.RemoveAll

    Dim i As Integer
    For i = list_series.ListCount - 1 To 0 Step -1
        list_series.RemoveItem (i)
    Next i

    Dim cht_obj As ChartObject

    Dim ser As series

    For Each cht_obj In Chart_GetObjectsFromObject(Selection)
        For Each ser In cht_obj.Chart.SeriesCollection

            Dim b_ser As bUTLChartSeries
            Set b_ser = New bUTLChartSeries

            b_ser.UpdateFromChartSeries ser

            Dim ser_name As Variant
            ser_name = IIf(Not b_ser.name Is Nothing, b_ser.name, "")

            list_series.AddItem
            If IsArray(ser_name) Then
                ser_name = ser_name(1, 1)
            End If

            list_series.List(list_series.ListCount - 1, 0) = ser_name


            list_series.List(list_series.ListCount - 1, 1) = b_ser.XValues.Address
            list_series.List(list_series.ListCount - 1, 2) = b_ser.Values.Address

            list_series.Selected(list_series.ListCount - 1) = True

            ser_coll.Add list_series.ListCount - 1 & ser_name, b_ser

        Next ser
    Next cht_obj
End Sub

Private Sub UserForm_Activate()
'load up the series from selection
    If dirty Then
        UpdateSeries
        dirty = False
    End If

    If list_series.ListCount = 0 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()

    dirty = True

End Sub

