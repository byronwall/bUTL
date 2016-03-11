VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_newCommands 
   Caption         =   "Additional Features"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   OleObjectBlob   =   "form_newCommands.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_newCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'---------------------------------------------------------------------------------------
' Module    : form_newCommands
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : This form is just buttons to easier get to new code
'---------------------------------------------------------------------------------------

Private Sub CommandButton1_Click()
    Chart_CreateDataLabels
End Sub

Private Sub CommandButton13_Click()
    ChartApplyToAll
End Sub

Private Sub CommandButton15_Click()
    Hide

    Dim frm As New form_chtSeries
    frm.Show
End Sub

Private Sub CommandButton16_Click()
    Selection_ColorWithHex
End Sub

Private Sub CommandButton17_Click()
    Chart_TrendlinesToAverage
End Sub

Private Sub CommandButton18_Click()
    ChartPropMove
End Sub

Private Sub CommandButton19_Click()
    Chart_RemoveTrendlines
End Sub

Private Sub CommandButton20_Click()
    PivotSetAllFields
End Sub

Private Sub CommandButton21_Click()
    ConvertSelectionToCsv
End Sub

Private Sub CommandButton22_Click()
    ColorInputs
End Sub

Private Sub CommandButton23_Click()
    UnhideAllRowsAndColumns
End Sub

Private Sub CommandButton25_Click()
    ExportFilesFromFolder
End Sub

Private Sub CommandButton26_Click()
    GenerateRandomData
End Sub

Private Sub CommandButton27_Click()
    SeriesSplitIntoBins
End Sub

Private Sub CommandButton28_Click()
    Chart_SortSeriesByName
End Sub

Private Sub CommandButton29_Click()
    CreatePdfOfEachXlsxFileInFolder
End Sub

Private Sub CommandButton30_Click()
    ApplyFormattingToEachColumn
End Sub

Private Sub CommandButton31_Click()
    ComputeDistanceMatrix
End Sub

Private Sub CommandButton32_Click()
    Chart_CreateChartWithSeriesForEachColumn
End Sub
