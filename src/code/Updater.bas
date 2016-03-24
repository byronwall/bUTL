Attribute VB_Name = "Updater"
Option Explicit

Public Sub UpdateSelf()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateSelf
    ' Author    : @byronwall
    ' Date      : 2016 02 05
    ' Purpose   : will download most recent version and replace current version of self with it
    '---------------------------------------------------------------------------------------
    '

    'TODO need to do a check at the top to determine the current version

    Dim str_url As String
    str_url = "http://github.com/byronwall/bUTL/releases/download/v0.1.1-beta/bUTL.xlam"

    Dim str_download_path As String
    str_download_path = ThisWorkbook.path & "\" & "butl-github.xlam"

    Download_File str_url, str_download_path

    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim Script As String, ScriptFile As String
    Dim A As AddIn
    Dim objShell


    ''actions to be done by the updater:

    'delete current file
    'rename downloaded file
    'message box that things are done
    'open excel back up
    'delete script

    Script = _
    "'sleep so Excel closes" & vbCrLf & _
             "Wscript.Sleep 1000" & vbCrLf & _
             "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
             "fso.DeleteFile ""[butl-current]""" & vbCrLf & _
             "fso.MoveFile ""[butl-new]"", ""[butl-current]""" & vbCrLf & _
             "Set objShell = CreateObject(""Wscript.Shell"")" & vbCrLf & _
             "objShell.Run ""excel.exe""" & vbCrLf & _
             "fso.DeleteFile Wscript.ScriptFullName" & vbCrLf & _
             "MsgBox ""bUTL is now updated to the current version"""

    Script = Replace(Script, "[butl-current]", ThisWorkbook.path & "\" & ThisWorkbook.name)
    Script = Replace(Script, "[butl-new]", str_download_path)

    
    Set fso = CreateObject("Scripting.FileSystemObject")

    ScriptFile = ThisWorkbook.path & "\" & "butl updater.vbs"
    Set ts = fso.CreateTextFile(ScriptFile)
    ts.Write Script
    ts.Close

    Dim str_scriptPath As String
    str_scriptPath = """" & ScriptFile & """"

    CreateObject("Wscript.Shell").Run str_scriptPath

    Application.DisplayAlerts = False
    Application.Quit
End Sub

