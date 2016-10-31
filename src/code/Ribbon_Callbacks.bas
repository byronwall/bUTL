Attribute VB_Name = "Ribbon_Callbacks"
Option Explicit

Dim targetForm_chartGrid As New form_chtGrid

Public Sub btn_aboutForm_onAction(control As IRibbonControl)

    'catch the rare case where the add-in is opened directly
    If ActiveWorkbook Is Nothing Then Application.Workbooks.Add

    ActiveWorkbook.FollowHyperlink "https://github.com/byronwall/bUTL"

End Sub

Sub btn_chartAddTitles_onAction(control As IRibbonControl)
    Chart_AddTitles
End Sub

Sub btn_chartApplyColors_onAction(control As IRibbonControl)
    Chart_ApplyTrendColors
End Sub

Sub btn_chartAxisTitleBySeries_onAction(control As IRibbonControl)
    Chart_AxisTitleIsSeriesTitle
End Sub

Public Sub btn_chartBothAxis_onAction(control As IRibbonControl)
    Chart_FitAxisToMaxAndMin xlCategory
    Chart_FitAxisToMaxAndMin xlValue
End Sub

Sub btn_chartExtendSeries_onAction(control As IRibbonControl)
    Chart_ExtendSeriesToRanges
End Sub

Public Sub btn_chartFindX_onAction(control As IRibbonControl)
    Chart_GoToXRange
End Sub

Public Sub btn_chartFindY_onAction(control As IRibbonControl)
    Chart_GoToYRange
End Sub

Sub btn_chartFitAutoX_onAction(control As IRibbonControl)
    Chart_Axis_AutoX
End Sub

Sub btn_chartFitAutoY_onAction(control As IRibbonControl)
    Chart_Axis_AutoY
End Sub

Public Sub btn_chartFitX_onAction(control As IRibbonControl)
    Chart_FitAxisToMaxAndMin xlCategory
End Sub

Public Sub btn_chartFlipXY_onAction(control As IRibbonControl)
    ChartFlipXYValues
End Sub

Public Sub btn_chartFormat_onAction(control As IRibbonControl)
    ChartDefaultFormat
End Sub

Public Sub btn_chartMergeSeries_onAction(control As IRibbonControl)
    ChartMergeSeries
End Sub

Public Sub btn_chartPivot_onAction(control As IRibbonControl)
    ChartDefaultFormat
End Sub

Sub btn_chartSplitSeries_onAction(control As IRibbonControl)
    ChartSplitSeries
End Sub

Sub btn_chartTimeSeries_onAction(control As IRibbonControl)
    CreateMultipleTimeSeries
End Sub

Sub btn_chartTrendLines_onAction(control As IRibbonControl)
    Chart_AddTrendlineToSeriesAndColor
End Sub

Public Sub btn_chartXYMatrix_onAction(control As IRibbonControl)
    ChartCreateXYGrid
End Sub

Public Sub btn_chartYAxis_onAction(control As IRibbonControl)
    Chart_FitAxisToMaxAndMin xlValue
End Sub

Public Sub btn_chtGrid_onAction(control As IRibbonControl)
    targetForm_chartGrid.Show
End Sub

Public Sub btn_colorCategory_onAction(control As IRibbonControl)
    CategoricalColoring
End Sub

Public Sub btn_colorize_onAction(control As IRibbonControl)
    Colorize
End Sub

Public Sub btn_convertValue_onAction(control As IRibbonControl)
    SelectedToValue
End Sub

Public Sub btn_copyClear_onAction(control As IRibbonControl)
    MsgBox "Copy clear is missing"
End Sub

Public Sub btn_cutTranspose_onAction(control As IRibbonControl)
    CutPasteTranspose
End Sub

Public Sub btn_extendArray_onAction(control As IRibbonControl)
    ExtendArrayFormulaDown
End Sub

Public Sub btn_fillDown_onAction(control As IRibbonControl)
    FillValueDown
End Sub

Sub btn_fmtDateTime_onAction(control As IRibbonControl)
    Selection.NumberFormat = "mm/dd/yyyy HH:MM"
End Sub

Public Sub btn_folder_onAction(control As IRibbonControl)
    OpenContainingFolder
End Sub

Public Sub btn_hyperlink_onAction(control As IRibbonControl)
    MakeHyperlinks
End Sub

Public Sub btn_joinCells_onAction(control As IRibbonControl)
    CombineCells
End Sub

Public Sub btn_openNewFeatures_onAction(control As IRibbonControl)
    Dim targetForm As New form_newCommands
    targetForm.Show
End Sub

Public Sub btn_panelCharts_onAction(control As IRibbonControl)
    MsgBox "feature is not implemented yet..."
End Sub

Public Sub btn_piRecalc_onAction(control As IRibbonControl)
    ForceRecalc
End Sub

Public Sub btn_protect_onAction(control As IRibbonControl)
    LockAllSheets
End Sub

Public Sub btn_rmvComments_onAction(control As IRibbonControl)
    MsgBox "RemoveComments missing"
End Sub

Public Sub btn_seriesSplit_onAction(control As IRibbonControl)
    SeriesSplit
End Sub

Sub btn_sheetDeleteHiddenRows_onAction(control As IRibbonControl)
    Sheet_DeleteHiddenRows
End Sub


Public Sub btn_sheetNamesOutput_onAction(control As IRibbonControl)
    OutputSheets
End Sub

Sub btn_sht_unhide_onAction(control As IRibbonControl)
    Dim targetSheet As Worksheet

    For Each targetSheet In Sheets
        targetSheet.Visible = xlSheetVisible
    Next targetSheet
End Sub

Public Sub btn_split_onAction(control As IRibbonControl)
    SplitAndKeep
End Sub

Public Sub btn_splitCol_onAction(control As IRibbonControl)
    SplitIntoColumns
End Sub

Public Sub btn_splitRows_onAction(control As IRibbonControl)
    SplitIntoRows
End Sub

Public Sub btn_toNumeric_onAction(control As IRibbonControl)
    ConvertToNumber
End Sub

Public Sub btn_trimSelection_onAction(control As IRibbonControl)
    TrimSelection
End Sub

Public Sub btn_unprotectAll_onAction(control As IRibbonControl)
    UnlockAllSheets
End Sub

Public Sub btn_updateScrollbars_onAction(control As IRibbonControl)
    UpdateScrollbars
End Sub

Public Sub btn_checkUpdates_onAction(control As IRibbonControl)
    CheckForUpdates
End Sub


Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    '---------------------------------------------------------------------------------------
    ' Procedure : RibbonOnLoad
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : OnLoad entry point for the add-in
    '---------------------------------------------------------------------------------------
    '
    SetUpKeyboardHooksForSelection

End Sub

