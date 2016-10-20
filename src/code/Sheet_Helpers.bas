Attribute VB_Name = "Sheet_Helpers"
Option Explicit


Public Sub LockAllSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : LockAllSheets
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Locks all sheets with the same password
    '---------------------------------------------------------------------------------------
    '
    Dim userPassword As Variant
    userPassword = Application.InputBox("Password to lock")

    If Not userPassword Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False

        'Changed to activeworkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim targetSheet As Worksheet
        For Each targetSheet In ActiveWorkbook.Sheets
            On Error Resume Next
            targetSheet.Protect (userPassword)
        Next

        Application.ScreenUpdating = True
    End If

End Sub


Public Sub OutputSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : OutputSheets
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Creates a new worksheet with a list and link to each sheet
    '---------------------------------------------------------------------------------------
    '
    Dim outputSheet As Worksheet
    Set outputSheet = Worksheets.Add(Before:=Worksheets(1))
    outputSheet.Activate

    Dim outputRange As Range
    Set outputRange = outputSheet.Range("B2")

    Dim targetRow As Long
    targetRow = 0

    Dim targetSheet As Worksheet
    For Each targetSheet In Worksheets

        If targetSheet.name <> outputSheet.name Then

            targetSheet.Hyperlinks.Add _
                outputRange.Offset(targetRow), "", _
                "'" & targetSheet.name & "'!A1", , _
                targetSheet.name
            targetRow = targetRow + 1

        End If
    Next targetSheet

End Sub


Public Sub UnlockAllSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : UnlockAllSheets
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Unlocks all sheets with the same password
    '---------------------------------------------------------------------------------------
    '
    Dim userPassword As Variant
    userPassword = Application.InputBox("Password to unlock")
    
    Dim errorCount As Long
    errorCount = 0
    
    If Not userPassword Then
        MsgBox "Cancelled."
    Else
        Application.ScreenUpdating = False
        'Changed to activeworkbook so if add-in is not installed, it will target the active book rather than the xlam
        Dim targetSheet As Worksheet
        For Each targetSheet In ActiveWorkbook.Sheets
            'Let's keep track of the errors to inform the user
            If Err.Number <> 0 Then errorCount = errorCount + 1
            Err.Clear
            On Error Resume Next
            targetSheet.Unprotect (userPassword)

        Next targetSheet
        If Err.Number <> 0 Then errorCount = errorCount + 1
        Application.ScreenUpdating = True
    End If
    If errorCount <> 0 Then
        MsgBox (errorCount & " sheets could not be unlocked due to bad password.")
    End If
End Sub


Public Sub AscendSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : AscendSheets
    ' Author    : @raymondwise
    ' Date      : 2015 08 07
    ' Purpose   : Places worksheets in ascending alphabetical order.
    '---------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    Dim countOfSheets As Long
    countOfSheets = targetWorkbook.Sheets.count

    Dim i As Long
    Dim j As Long

    With targetWorkbook
        For j = 1 To countOfSheets
            For i = 1 To countOfSheets - 1
                If UCase(.Sheets(i).name) > UCase(.Sheets(i + 1).name) Then .Sheets(i).Move after:=.Sheets(i + 1)
            Next i
        Next j
    End With

    Application.ScreenUpdating = True
End Sub

Public Sub DescendSheets()
    '---------------------------------------------------------------------------------------
    ' Procedure : DescendSheets
    ' Author    : @raymondwise
    ' Date      : 2015 08 07
    ' Purpose   : Places worksheets in descending alphabetical order.
    '---------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ActiveWorkbook

    Dim countOfSheets As Long
    countOfSheets = targetWorkbook.Sheets.count

    Dim i As Long
    Dim j As Long

    With targetWorkbook
        For j = 1 To countOfSheets
            For i = 1 To countOfSheets - 1
                If UCase(.Sheets(i).name) < UCase(.Sheets(i + 1).name) Then .Sheets(i).Move after:=.Sheets(i + 1)
            Next i
        Next j
    End With

    Application.ScreenUpdating = True
End Sub

