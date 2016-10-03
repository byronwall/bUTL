Attribute VB_Name = "Updater"
Option Explicit

Public Sub CheckForUpdates()
    
    'get the path of the current file
    Dim strReleaseUrl As String
    strReleaseUrl = "https://api.github.com/repos/byronwall/butl/releases"
    
    Dim strGitHubData As String
    strGitHubData = DownloadFileAsString(strReleaseUrl)
    
    'this will grab the first file from the most recent release
    'this is a cheap way to "parse" the JSON without a library
    Dim strUrl As String
    strUrl = Split(Split(strGitHubData, "tag_name"":")(1), """")(1)
    
    Dim strVersion As String
    strVersion = "Current version on GitHub is " & vbCrLf & _
                 vbTab & strUrl & vbCrLf & _
                 "Version of bUTL on computer is" & vbCrLf & _
                 vbTab & bUTL_GetVersion() & vbCrLf & _
                 "Do you want to update?"
    
    Dim shouldUpdate As VbMsgBoxResult
    shouldUpdate = MsgBox(strVersion, vbYesNo, "Update?")
    
    If shouldUpdate = vbYes Then
        UpdateSelf
    End If

End Sub

Public Sub UpdateSelf()
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateSelf
    ' Author    : @byronwall
    ' Date      : 2016 02 05
    ' Purpose   : will download most recent version and replace current version of self with it
    '---------------------------------------------------------------------------------------
    '
    
    Dim promptResults As VbMsgBoxResult
    promptResults = MsgBox("This will: download the latest bUTL file, close Excel, " & _
                           "create a VB script file to copy that file over the current one, and reopen Excel." & vbCrLf & vbCrLf & _
                           "Continue?", _
                           vbYesNo, "Update bUTL?")
                           
    If promptResults = vbNo Then
        Exit Sub
    End If
    
    'get the path of the current file
    Dim strReleaseUrl As String
    strReleaseUrl = "https://api.github.com/repos/byronwall/butl/releases"
    
    Dim strGitHubData As String
    strGitHubData = DownloadFileAsString(strReleaseUrl)
    
    'this will grab the first file from the most recent release
    'this is a cheap way to "parse" the JSON without a library
    Dim strUrl As String
    strUrl = Split(Split(strGitHubData, "browser_download_url"":")(1), """")(1)
    
    Debug.Print strUrl

    Dim str_download_path As String
    str_download_path = ThisWorkbook.path & "\" & "butl-github.xlam"

    Download_File strUrl, str_download_path

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

