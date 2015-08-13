Attribute VB_Name = "SelectionMgr"
'---------------------------------------------------------------------------------------
' Module    : SelectionMgr
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : This module contains code related to changing the Selection with kbd shortcuts
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : OffsetSelectionByRowsAndColumns
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Offsets and selects the Selection a given number of rows/columns
'---------------------------------------------------------------------------------------
'
Sub OffsetSelectionByRowsAndColumns(iRowsOff As Integer, iColsOff As Integer)

    If TypeOf Selection Is Range Then

        'this error should only get called if the new range is outside the sheet boundaries
        On Error GoTo OffsetSelectionByRowsAndColumns_Exit

        Selection.Offset(iRowsOff, iColsOff).Select

        On Error GoTo 0
    End If

OffsetSelectionByRowsAndColumns_Exit:

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectionOffsetDown
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Moves Selection down one row
'---------------------------------------------------------------------------------------
'
Sub SelectionOffsetDown()

    Call OffsetSelectionByRowsAndColumns(1, 0)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectionOffsetLeft
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Moves Selection left one column
'---------------------------------------------------------------------------------------
'
Sub SelectionOffsetLeft()

    Call OffsetSelectionByRowsAndColumns(0, -1)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectionOffsetRight
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Moves selection right one column
'---------------------------------------------------------------------------------------
'
Sub SelectionOffsetRight()

    Call OffsetSelectionByRowsAndColumns(0, 1)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectionOffsetUp
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Moves Selection up one row
'---------------------------------------------------------------------------------------
'
Sub SelectionOffsetUp()

    Call OffsetSelectionByRowsAndColumns(-1, 0)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetUpKeyboardHooksForSelection
' Author    : @byronwall
' Date      : 2015 08 05
' Purpose   : Creates hotkey events for the selection events
'---------------------------------------------------------------------------------------
'
Sub SetUpKeyboardHooksForSelection()

    Application.OnKey "^%{RIGHT}", "SelectionOffsetRight"
    Application.OnKey "^%{LEFT}", "SelectionOffsetLeft"
    Application.OnKey "^%{UP}", "SelectionOffsetUp"
    Application.OnKey "^%{DOWN}", "SelectionOffsetDown"

End Sub

