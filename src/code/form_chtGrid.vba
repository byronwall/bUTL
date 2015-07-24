VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_chtGrid 
   Caption         =   "Chart Grid"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1755
   OleObjectBlob   =   "form_chtGrid.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_chtGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub btn_grid_Click()
    
    Chart_GridOfCharts _
        txt_cols, _
        CDbl(txt_width), _
        CDbl(txt_height), _
        CDbl(txt_vOff), _
        CDbl(txt_hOff), _
        chk_down, _
        chk_zoom
    
    Hide
    
End Sub
