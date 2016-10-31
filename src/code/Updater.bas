Attribute VB_Name = "Updater"
Option Explicit

Public Sub CheckForUpdates()
    
    'get the path of the current file
    Const RELEASE_URL As String = "https://api.github.com/repos/byronwall/butl/releases"
    
    Dim githubData As String
    githubData = DownloadFileAsString(RELEASE_URL)
    
    'this will grab the first file from the most recent release
    'this is a cheap way to "parse" the JSON without a library
    Dim splitURL As String
    splitURL = Split(Split(githubData, "tag_name"":")(1), """")(1)
    
    Dim currentVersion As String
    currentVersion = "Current version on GitHub is " & vbCrLf & _
                 vbTab & splitURL & vbCrLf & _
                 "Version of bUTL on computer is" & vbCrLf & _
                 vbTab & bUTL_GetVersion() & vbCrLf & _
                 "Do you want to update?"
    
    Dim shouldUpdate As VbMsgBoxResult
    shouldUpdate = MsgBox(currentVersion, vbYesNo, "Update?")
    
    If shouldUpdate = vbYes Then UpdateSelf


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
                           
    If promptResults = vbNo Then Exit Sub

    
    'get the path of the current file
    Const RELEASE_URL As String = "https://api.github.com/repos/byronwall/butl/releases"
    
    Dim githubData As String
    githubData = DownloadFileAsString(RELEASE_URL)
    
    'this will grab the first file from the most recent release
    'this is a cheap way to "parse" the JSON without a library
    Dim splitURL As String
    splitURL = Split(Split(githubData, "browser_download_url"":")(1), """")(1)
    
    Debug.Print splitURL

    Dim downloadPath As String
    downloadPath = ThisWorkbook.path & "\" & "butl-github.xlam"

    Download_File splitURL, downloadPath

    Dim fso As FileSystemObject
    Dim textStream As textStream
    Dim scriptAddress As String, ScriptFile As String
    Dim A As AddIn
    'objShell should have a type
    Dim objShell


    ''actions to be done by the updater:

    'delete current file
    'rename downloaded file
    'message box that things are done
    'open excel back up
    'delete script

    scriptAddress = _
           "'sleep so Excel closes" & vbCrLf & _
           "Wscript.Sleep 1000" & vbCrLf & _
           "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
           "fso.DeleteFile ""[butl-current]""" & vbCrLf & _
           "fso.MoveFile ""[butl-new]"", ""[butl-current]""" & vbCrLf & _
           "Set objShell = CreateObject(""Wscript.Shell"")" & vbCrLf & _
           "objShell.Run ""excel.exe""" & vbCrLf & _
           "fso.DeleteFile Wscript.ScriptFullName" & vbCrLf & _
           "MsgBox ""bUTL is now updated to the current version"""

    scriptAddress = Replace(scriptAddress, "[butl-current]", ThisWorkbook.path & "\" & ThisWorkbook.name)
    scriptAddress = Replace(scriptAddress, "[butl-new]", downloadPath)

    
    Set fso = CreateObject("Scripting.FileSystemObject")

    ScriptFile = ThisWorkbook.path & "\" & "butl updater.vbs"
    Set textStream = fso.CreateTextFile(ScriptFile)
    textStream.Write scriptAddress
    textStream.Close

    Dim pathToScript As String
    pathToScript = """" & ScriptFile & """"

    CreateObject("Wscript.Shell").Run pathToScript

    Application.DisplayAlerts = False
    Application.Quit
End Sub

