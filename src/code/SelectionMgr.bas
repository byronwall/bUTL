Attribute VB_Name = "SelectionMgr"
Option Explicit


Public Sub OffsetSelectionByRowsAndColumns(ByVal numberOfRows As Long, ByVal numberOfColumns As Long)
    '---------------------------------------------------------------------------------------
    ' Procedure : OffsetSelectionByRowsAndColumns
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : Offsets and selects the Selection a given number of rows/columns
    '---------------------------------------------------------------------------------------
    '
    If TypeOf Selection Is Range Then

        'this error should only get called if the new range is outside the sheet boundaries
        On Error GoTo OffsetSelectionByRowsAndColumns_Exit

        Selection.Offset(numberOfRows, numberOfColumns).Select

        On Error GoTo 0
    End If

OffsetSelectionByRowsAndColumns_Exit:

End Sub


Public Sub SelectionOffsetDown()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectionOffsetDown
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : Moves Selection down one row
    '---------------------------------------------------------------------------------------
    '
     OffsetSelectionByRowsAndColumns 1, 0

End Sub


Public Sub SelectionOffsetLeft()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectionOffsetLeft
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : Moves Selection left one column
    '---------------------------------------------------------------------------------------
    '
     OffsetSelectionByRowsAndColumns 0, -1

End Sub


Public Sub SelectionOffsetRight()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectionOffsetRight
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : Moves selection right one column
    '---------------------------------------------------------------------------------------
    '
    OffsetSelectionByRowsAndColumns 0, 1

End Sub


Public Sub SelectionOffsetUp()
    '---------------------------------------------------------------------------------------
    ' Procedure : SelectionOffsetUp
    ' Author    : @byronwall
    ' Date      : 2015 08 05
    ' Purpose   : Moves Selection up one row
    '---------------------------------------------------------------------------------------
    '
     OffsetSelectionByRowsAndColumns -1, 0

End Sub


Public Sub SetUpKeyboardHooksForSelection()
    '---------------------------------------------------------------------------------------
    ' Procedure : SetUpKeyboardHooksForSelection
    ' Author    : @byronwall
    ' Date      : 2016 09 29
    ' Purpose   : Creates hotkey events for the selection events
    '---------------------------------------------------------------------------------------
    '
    
    'SHIFT =    +
    'CTRL =     ^
    'ALT =      %

    'set up the keys for the selection mover
    Application.OnKey "^%{RIGHT}", "SelectionOffsetRight"
    Application.OnKey "^%{LEFT}", "SelectionOffsetLeft"
    Application.OnKey "^%{UP}", "SelectionOffsetUp"
    Application.OnKey "^%{DOWN}", "SelectionOffsetDown"
    
    'set up the keys for the indent level
    Application.OnKey "+^%{RIGHT}", "Formatting_IncreaseIndentLevel"
    Application.OnKey "+^%{LEFT}", "Formatting_DecreaseIndentLevel"

End Sub

