Attribute VB_Name = "Sheet_Helpers"
Option Explicit

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

        'Changed to activeworkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim sheet As Worksheet
        For Each sheet In ActiveWorkbook.Sheets
            On Error Resume Next
            sheet.Protect (pass)
        Next

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

    Dim newSheet As Worksheet
    Set newSheet = Worksheets.Add(Before:=Worksheets(1))
    newSheet.Activate

    Dim newRange As Range
    Set newRange = newSheet.Range("B2")

    Dim row As Long
    row = 0

    Dim mySheet As Worksheet
    For Each mySheet In Worksheets

        If mySheet.name <> newSheet.name Then

            mySheet.Hyperlinks.Add _
                    newRange.Offset(row), "", _
                    "'" & mySheet.name & "'!A1", , _
                        mySheet.name
            row = row + 1

        End If
    Next mySheet

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
    
    Dim iErr As Long
    iErr = 0
    
    If pass = False Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False
        'Changed to activeworkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim mySheet As Worksheet
        For Each mySheet In ActiveWorkbook.Sheets
            'Let's keep track of the errors to inform the user
            If Err.Number <> 0 Then iErr = iErr + 1
            Err.Clear
            On Error Resume Next
            mySheet.Unprotect (pass)

        Next mySheet
        If Err.Number <> 0 Then iErr = iErr + 1
        Application.ScreenUpdating = True
    End If
    If iErr <> 0 Then
    MsgBox (iErr & " sheets could not be unlocked due to bad password.")
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AscendSheets
' Author    : @raymondwise
' Date      : 2015 08 07
' Purpose   : Places worksheets in ascending alphabetical order.
'---------------------------------------------------------------------------------------
Sub AscendSheets()
Application.ScreenUpdating = False
Dim myBook As Workbook
Set myBook = ActiveWorkbook

Dim numberOfSheets As Long
numberOfSheets = myBook.Sheets.count

Dim i As Long
Dim j As Long

With myBook
    For j = 1 To numberOfSheets
        For i = 1 To numberOfSheets - 1
            If UCase(.Sheets(i).name) > UCase(.Sheets(i + 1).name) Then
                .Sheets(i).Move after:=.Sheets(i + 1)
            End If
        Next i
    Next j
End With

Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DescendSheets
' Author    : @raymondwise
' Date      : 2015 08 07
' Purpose   : Places worksheets in descending alphabetical order.
'---------------------------------------------------------------------------------------
Sub DescendSheets()
Application.ScreenUpdating = False
Dim myBook As Workbook
Set myBook = ActiveWorkbook

Dim numberOfSheets As Long
numberOfSheets = myBook.Sheets.count

Dim i As Long
Dim j As Long

With myBook
    For j = 1 To numberOfSheets
        For i = 1 To numberOfSheets - 1
            If UCase(.Sheets(i).name) < UCase(.Sheets(i + 1).name) Then
                .Sheets(i).Move after:=.Sheets(i + 1)
            End If
        Next i
    Next j
End With

Application.ScreenUpdating = True
End Sub

