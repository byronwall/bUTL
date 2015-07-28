Attribute VB_Name = "BuildFile"
Option Explicit

Sub CreateFileFromPackageAndCode()

'workbook will run in /scripts/, files are in /src/, output to /

    Dim FSO As New FileSystemObject
    Dim folRoot As Folder
    Set folRoot = FSO.GetFolder(ThisWorkbook.Path).ParentFolder

    Dim strZip As Variant
    strZip = FSO.BuildPath(folRoot, "temp.zip")

    Dim strPackage As Variant
    strPackage = FSO.BuildPath(folRoot, "src\package\")

    Dim strPath As String
    strPath = FSO.BuildPath(folRoot, "src\code\")

    Dim strAddIn As String
    strAddIn = Replace(strZip, "temp.zip", "bUTL.xlam")
    
    'delete add-in if it exists
    If FSO.FileExists(strAddIn) Then
        FSO.DeleteFile strAddIn
    End If

    NewZip strZip

    Dim oApp As Object
    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace(strZip).CopyHere oApp.Namespace(strPackage).items

    'Keep script waiting until Compressing is done
    Do Until oApp.Namespace(strZip).items.Count = _
       oApp.Namespace(strPackage).items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop

    'convert zip to xlam
    Name strZip As strAddIn

    'open add-in
    Dim wkAddIn As Workbook
    Set wkAddIn = Application.Workbooks.Open(strAddIn)

    'remove all existing code (except for Sheet1 and ThisWorkbook
    'requires a reference to the VBA Extensibility file
    With wkAddIn.VBProject
        Dim i As Integer
        For i = .VBComponents.Count To 1 Step -1
            If .VBComponents(i).Type <> vbext_ct_Document Then
                'Static folder name
                .VBComponents.Remove .VBComponents(i)
            End If
        Next i
    End With

    'iterate through code files and add them in
    Dim strFile As Variant
    strFile = Dir(strPath)

    While (strFile <> "")
        If InStr(strFile, ".vba") > 0 Then
            wkAddIn.VBProject.VBComponents.Import FSO.BuildPath(strPath, strFile)
        End If
        strFile = Dir
    Wend

    wkAddIn.Save
    wkAddIn.Close

End Sub

Sub CreatePackageAndCodeFromFile()


    Dim FSO As New FileSystemObject
    Dim folRoot As Folder
    Set folRoot = FSO.GetFolder(ThisWorkbook.Path).ParentFolder

    Dim strZip As Variant
    strZip = FSO.BuildPath(folRoot, "temp.zip")

    Dim strPackage As Variant
    strPackage = FSO.BuildPath(folRoot, "src\package")

    Dim strPath As String
    strPath = FSO.BuildPath(folRoot, "src\code")

    Dim strAddIn As String
    strAddIn = FSO.BuildPath(folRoot, "bUTL.xlam")

    'remove all existing files
    If FSO.FolderExists(strPackage) Then
        FSO.DeleteFolder strPackage
    End If

    FSO.CreateFolder strPackage

    If FSO.FolderExists(strPath) Then
        FSO.DeleteFolder strPath
    End If

    FSO.CreateFolder strPath

    'open workbook
    Dim wkAddIn As Workbook
    Set wkAddIn = Application.Workbooks.Open(strAddIn)

    'export all existing code (except for Sheet1 and ThisWorkbook
    'requires a reference to the VBA Extensibility file
    With wkAddIn.VBProject
        Dim i As Integer
        For i = .VBComponents.Count To 1 Step -1
            If .VBComponents(i).Type <> vbext_ct_Document Then
                .VBComponents(i).Export FSO.BuildPath(strPath, .VBComponents(i).CodeModule.Name & ".vba")
            End If
        Next i
    End With

    'close butl so it can be unzipped
    wkAddIn.Close

    Name strAddIn As strZip

    Dim oApp As Object
    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.Namespace(strPackage).CopyHere oApp.Namespace(strZip).items

    'rename it back
    Name strZip As strAddIn

End Sub

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

