Attribute VB_Name = "Sheet_Helpers"
'---------------------------------------------------------------------------------------
' Module    : Sheet_Helpers
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Contains code related to sheets and sheet processing
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : LockAllSheets
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Locks all sheets with the same password
'---------------------------------------------------------------------------------------
'
Sub LockAllSheets()

    Dim pass As Variant
    pass = Application.InputBox("Password to lock")

    If pass = False Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False

        Dim sht As Worksheet
        For Each sht In ThisWorkbook.Sheets
            On Error Resume Next
            sht.Protect pass
        Next sht

        Application.ScreenUpdating = True
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : OutputSheets
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Creates a new worksheet with a list and link to each sheet
'---------------------------------------------------------------------------------------
'
Sub OutputSheets()

    Dim sht_out As Worksheet
    Set sht_out = Worksheets.Add(Before:=Worksheets(1))
    sht_out.Activate

    Dim rng_out As Range
    Set rng_out = sht_out.Range("B2")

    Dim int_row As Integer
    int_row = 0

    Dim sht As Worksheet
    For Each sht In Worksheets

        If sht.name <> sht_out.name Then

            sht.Hyperlinks.Add _
                    rng_out.Offset(int_row), "", _
                    "'" & sht.name & "'!A1", , _
                        sht.name
            int_row = int_row + 1

        End If
    Next sht

End Sub

'---------------------------------------------------------------------------------------
' Procedure : UnlockAllSheets
' Author    : @byronwall
' Date      : 2015 07 24
' Purpose   : Unlocks all sheets with the same password
'---------------------------------------------------------------------------------------
'
Sub UnlockAllSheets()

    Dim pass As Variant
    pass = Application.InputBox("Password to unlock")

    If pass = False Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False

        Dim sht As Worksheet
        For Each sht In ThisWorkbook.Sheets
            On Error Resume Next
            sht.Unprotect pass

        Next sht

        Application.ScreenUpdating = True
    End If

End Sub

