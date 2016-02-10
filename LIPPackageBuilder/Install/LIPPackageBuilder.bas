Attribute VB_Name = "LIPPackageBuilder"
Option Explicit

Private m_TemporaryFolder As String

Public Sub OpenPackageBuilder()
    On Error GoTo ErrorHandler
    Dim oDialog As New Lime.Dialog
    Dim idpersons As String
    Dim oItem As Lime.ExplorerItem
    oDialog.Type = lkDialogHTML
    oDialog.Property("url") = Application.WebFolder & "lbs.html?ap=packagebuilder&type=tab"
    oDialog.Property("height") = 900
    oDialog.Property("width") = 1600
    oDialog.show

    Exit Sub
ErrorHandler:
    Call UI.ShowError("Globals.OpenPackageBuilder")
End Sub


Public Function GetVBAComponents() As String
On Error GoTo ErrorHandler
    Dim oComp As VBComponent
    Dim strComponents As String
    strComponents = "["
    For Each oComp In Application.VBE.ActiveVBProject.VBComponents
        'Only include modules, class modules and forms
        If oComp.Type <> vbext_ct_ActiveXDesigner And oComp.Type <> vbext_ct_Document Then
            strComponents = strComponents & "{"
            strComponents = strComponents & """name"": """ & oComp.Name & ""","
            strComponents = strComponents & """type"": """ & GetModuleTypeName(oComp.Type) & """},"
            
        End If
    Next
    
    strComponents = VBA.Left(strComponents, Len(strComponents) - 1)
    strComponents = strComponents + "]"
    
    GetVBAComponents = strComponents
Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.GetVBAComponents")
End Function

Private Function GetModuleTypeName(ModuleType As Long) As String
On Error GoTo ErrorHandler
    Dim strModuleTypeName As String
    strModuleTypeName = ""
    Select Case ModuleType
        Case 1:
            strModuleTypeName = "Module"
        Case 2:
            strModuleTypeName = "Class Module"
        Case 3:
            strModuleTypeName = "Form"
        Case Else
            strModuleTypeName = "Other"
    End Select
    GetModuleTypeName = strModuleTypeName
Exit Function
ErrorHandler:
Call UI.ShowError("LIPPackageBuilder.GetModuleTypeName")
End Function


Public Sub CreatePackage(strPackageJson As String)
On Error GoTo ErrorHandler
    Dim strTempFolder As String
    Dim oPackage As Object
    Dim bResult As Boolean
    bResult = True
    'Create temporary folder
    strTempFolder = CreateTemporaryFolder()
    Set oPackage = JsonConverter.ParseJson(strPackageJson)
    
    'Export VBA modules
    If bResult Then
        bResult = ExportVBA(oPackage)
    End If
    If bResult = False Then
        Call Application.MessageBox("Couldn't export VBA Modules.")
        Exit Sub
    End If
    
    'Save Package.json
    If bResult Then
        bResult = SavePackageFile(oPackage, strTempFolder)
    End If
    If bResult = False Then
        Call Application.MessageBox("Couldn't save the package.json file.")
        Exit Sub
    End If
    'Rename Temporary folder
    If bResult Then
        bResult = RenameTemporaryFolder(oPackage, strTempFolder)
    End If
    
    If bResult = False Then
        Call Application.MessageBox("Couldn't Rename the temporary folder.")
        Exit Sub
    End If
    
    'Zip Temporary folder and save package
    Dim ZipPath As String
    If bResult Then
        bResult = ZipTemporaryFolder(oPackage.Item("name"), strTempFolder, ZipPath)
    End If
    
    If bResult = False Then
        Call Application.MessageBox("Couldn't save the package Zip file")
        Exit Sub
    End If
    
    Call Application.Shell(ZipPath)
    
    'Delete Temporary folder
    If bResult Then
        bResult = DeleteTemporaryFolder(strTempFolder)
    End If
    
    
Exit Sub
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.CreatePackage")
End Sub

Public Function GetFolder() As String
On Error GoTo ErrorHandler
    Dim fldr As New LCO.FolderDialog
    Dim sItem As String
    
    GetFolder = ""
        
    fldr.Text = "Select a Folder to save the package file."
    If fldr.show = vbOK Then
        GetFolder = fldr.Folder
    End If
    Exit Function
ErrorHandler:
    GetFolder = ""
    Set fldr = Nothing
End Function

Private Function ZipTemporaryFolder(strPackageName As String, strTempFolder As String, ByRef ZipPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim FileNameZip, FolderName
    Dim strDate As String, DefPath As String
    Dim oApp As Object
    Dim bResult As Boolean
    bResult = True
    DefPath = GetFolder()
    If DefPath = "" Then
        ZipTemporaryFolder = False
        Exit Function
    End If
    
    ZipPath = DefPath
    'Make sure the path format is as it's expected by the NewZip function
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    

    FileNameZip = DefPath & strPackageName & ".zip"

    'Create empty Zip File
    Call NewZip(FileNameZip)
    Dim oZipFile As Object
    Dim oPackageFolder As Object
    Set oApp = CreateObject("Shell.Application")
    'Create folder object for the zip file
    Set oZipFile = oApp.NameSpace(FileNameZip)
    
    If Not oZipFile Is Nothing Then
        
        
        'Create folder object for the package folder (different path format, which is messed up...)
        Set oPackageFolder = oApp.NameSpace(strTempFolder & "\")
        If Not oPackageFolder Is Nothing Then
            'Move files from the package folder to the zip file
            oZipFile.CopyHere oPackageFolder.Items
        
            'Keep script waiting until Compressing is done
            On Error Resume Next
            Do Until oZipFile.Items.Count = _
               oPackageFolder.Items.Count
                Application.Wait (Now + TimeValue("0:00:01"))
            Loop
            On Error GoTo 0
        Else
            FileNameZip = ""
            bResult = False
        End If
    Else
        FileNameZip = ""
        bResult = False
    End If
    ZipTemporaryFolder = bResult
Exit Function
ErrorHandler:
    ZipTemporaryFolder = False
    
End Function

Private Function RenameTemporaryFolder(oPackage As Object, strTempFolder As String) As Boolean
On Error GoTo ErrorHandler
    Dim bResult As Boolean
    bResult = True
    'I am assuming that the Folder Exists

    Dim NewFolderName As String
    'Name the temporary folder the same as the Package name
    If Right(strTempFolder, 1) = "\" Then
        NewFolderName = Left(strTempFolder, Len(strTempFolder) - 1)
    Else
        NewFolderName = strTempFolder
    End If
    
    NewFolderName = VBA.Left(NewFolderName, InStrRev(NewFolderName, "\")) & oPackage.Item("name")
    
    '-- Rename them
    Name strTempFolder As NewFolderName
    
    strTempFolder = NewFolderName

    RenameTemporaryFolder = bResult
Exit Function
ErrorHandler:
    bResult = False
End Function

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Function DeleteTemporaryFolder(strTempFolder As String) As Boolean
On Error GoTo ErrorHandler

    'Delete all files and subfolders
    'Be sure that no file is open in the folder
    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Right(strTempFolder, 1) = "\" Then
        strTempFolder = Left(strTempFolder, Len(strTempFolder) - 1)
    End If

    If FSO.FolderExists(strTempFolder) = False Then
        Exit Function
    End If

    On Error Resume Next
    'Delete files
    FSO.DeleteFile strTempFolder & "\*.*", True
    'Delete subfolders
    FSO.DeleteFolder strTempFolder & "\*.*", True
    On Error GoTo 0
    
    DeleteTemporaryFolder = True
    
    Exit Function
ErrorHandler:
    DeleteTemporaryFolder = False
End Function

Public Function SavePackageFile(oPackage As Object, strTempPath As String) As Boolean
On Error GoTo ErrorHandler
    Dim bResult As Boolean
    Dim FSO As New FileSystemObject
    Dim filePath As String
    filePath = strTempPath & "\package.json"
    bResult = True
    'Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(filePath, True, True)
    'Convert to a string and save
    Call oFile.WriteLine(JsonConverter.ConvertToJson(oPackage))
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing
    
    
    SavePackageFile = bResult
Exit Function
ErrorHandler:
    bResult = False
End Function


'Exports all VBA-Modules marked in the Package JSON
Public Function ExportVBA(oPackage As Object) As Boolean
On Error GoTo ErrorHandler
    Dim bResult As Boolean
    bResult = True
    If Not oPackage.Item("install") Is Nothing Then
        Dim oModule As Object
        
        If Not oPackage.Item("install").Item("vba") Is Nothing Then
            For Each oModule In oPackage.Item("install").Item("vba")
                bResult = ExportVBAModule(oModule.Item("name"))
                If bResult = False Then
                    ExportVBA = False
                    Exit Function
                End If
            Next
        End If
    End If
    ExportVBA = bResult
Exit Function
ErrorHandler:
    bResult = False
End Function

'Exporterar alla VBA-objekt till fil
Public Function ExportVBAModule(ModuleName As String) As Boolean
On Error GoTo ErrorHandler
    Dim Component As VBIDE.VBComponent
    Dim strInstallFolder As String
    Dim strTempFolder As String
    Dim bResult As Boolean
    bResult = True
    Set Component = ThisApplication.VBE.ActiveVBProject.VBComponents(ModuleName)
    If strTempFolder = "" Then
        strTempFolder = CreateTemporaryFolder()
        strInstallFolder = CreateTemporaryFolder("Install")
    End If
    Dim strFileName As String
    
    If Not Component Is Nothing Then
        strFileName = Component.Name
        Select Case Component.Type
            Case 1
                strFileName = strFileName & ".bas"
            Case 2
                strFileName = strFileName & ".cls"
            Case 3
                strFileName = strFileName & ".frm"
            
            Case Else
                bResult = False
                Exit Function
        End Select
        
        Call Component.Export(strInstallFolder & "\" & strFileName)
        bResult = True
    End If
    ExportVBAModule = bResult
Exit Function
ErrorHandler:
    bResult = False
End Function

Private Function CreateTemporaryFolder(Optional Subfolder As String = "") As String
On Error GoTo ErrorHandler
    'Kolla om sökvägen finns och skapar mappen
    Dim strTempPath As String
    strTempPath = Application.WebFolder & "apps\LIPPackageBuilder\LIPTemp"
    
    Dim strExists As String
    strExists = VBA.Dir(WebFolder & "apps\LIPPackageBuilder\LIPTemp", vbDirectory)
    If strExists = "" Then
        Call MkDir(strTempPath)
    End If
    
    If Subfolder <> "" Then
        strTempPath = WebFolder & "apps\LIPPackageBuilder\LIPTemp\" & Subfolder
        strExists = VBA.Dir(WebFolder & "apps\LIPPackageBuilder\LIPTemp\" & Subfolder, vbDirectory)
        If strExists = "" Then
            Call MkDir(strTempPath)
        End If
    End If
    
    CreateTemporaryFolder = strTempPath
    
Exit Function
ErrorHandler:
    Call UI.ShowError("LIPPackageBuilder.CreateTemporaryFolder")
End Function
