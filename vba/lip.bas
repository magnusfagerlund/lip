Attribute VB_Name = "lip"
Option Explicit

'Lundalogik Package Store, DO NOT CHANGE, used to download system files for LIP
'Please add your own stores in packages.json
Private Const BaseURL As String = "http://api.lime-bootstrap.com"
Private Const PackageStoreApiURL As String = "/packages/"
Private Const AppStoreApiURL As String = "/apps/"

Private Const DefaultInstallPath = "packages\"

Private IndentLenght As String
Private Indent As String
Private sLog As String

Public Sub UpgradePackage(Optional PackageName As String, Optional Path As String)
On Error GoTo ErrorHandler:
    If PackageName = "" Then
        'Upgrade all packages
        Call InstallFromPackageFile
    Else
        'Upgrade specific package
        Call Install(PackageName, True)
    End If
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.UpgradePackage")
End Sub

'Install package/app. Selects packagestore from packages.json

Public Sub Install(PackageName As String, Optional upgrade As Boolean, Optional Simulate As Boolean = True)
On Error GoTo ErrorHandler
    Dim Package As Object
    Dim PackageVersion As Double
    Dim downloadURL As String
    Dim InstallPath As String
    Dim bOk As Boolean
    Dim bLocalPackage As Boolean
    
    Application.MousePointer = 11

    IndentLenght = "  "
    sLog = ""
    bOk = True

    'Check if first use ever
    If Dir(WebFolder + "packages.json") = "" Then
        sLog = sLog + Indent + "No packages.json found, assuming fresh install" + vbNewLine
        Call InstallLIP
    End If
    
    'TODO Check if LIP has a new version
    Debug.Print Indent + "Updating LIP if necessary"
    Call UpdateLIPOnNewVersion
        
    PackageName = PackageName

    sLog = sLog + Indent + "====== LIP Install: " + PackageName + " ======" + vbNewLine

    sLog = sLog + Indent + "Looking for package: '" + PackageName + "'" + vbNewLine
    Set Package = SearchForPackageInStores(PackageName)
    
    If Package Is Nothing Then
        Application.MousePointer = 0
        Exit Sub
    End If
    
    
    If Package.Exists("source") Then
        downloadURL = VBA.Replace(Package.Item("source"), "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with URLs
    Else
        'Handle local source
        If Package.Exists("localsource") Then
            downloadURL = Package.Item("localsource")
            Call InstallFromZip(downloadURL)
            Exit Sub
        Else
            downloadURL = BaseURL & PackageStoreApiURL & PackageName & "/download/"  'Use Lundalogik Packagestore if source-node wasn't found
        End If
        
    End If

    If Package.Exists("installPath") Then
        InstallPath = ThisApplication.WebFolder & Package.Item("installPath") & "\"
    Else
        InstallPath = ThisApplication.WebFolder & DefaultInstallPath
    End If

    Set Package = Package

    'Parse result from store
    PackageVersion = findNewestVersion(Package.Item("versions"))

    'Check if package already exsists
    If Not upgrade Then
        If CheckForLocalInstalledPackage(PackageName, PackageVersion) = True Then
            Call Lime.MessageBox("Package already installed. If you want to upgrade the package, run command: " & vbNewLine & vbNewLine & "Call lip.Install(""" & PackageName & """, True)", vbInformation)
            Exit Sub
        End If
    End If

    'Install dependecies
    If Package.Exists("dependencies") Then
        IncreaseIndent
        Call InstallDependencies(Package, Simulate)
        DecreaseIndent
    End If

    'Download and unzip
    sLog = sLog + Indent + "Downloading '" + PackageName + "' files..." + vbNewLine
    Dim strDownloadError As String
    strDownloadError = DownloadFile(PackageName, downloadURL, InstallPath)
    If strDownloadError = "" Then
        Call Unzip(PackageName, InstallPath)
        sLog = sLog + Indent + "Download complete!" + vbNewLine
    
        If InstallPackageComponents(PackageName, PackageVersion, Package, InstallPath, Simulate) = False Then
            bOk = False
        End If
    Else
        bOk = False
        sLog = sLog + Indent + "Error: Could not download " + PackageName + " from url: " + downloadURL
    End If
    
    If bOk Then
        If Simulate Then
            sLog = sLog + Indent + "Simulation of " + PackageName + " done!" + vbNewLine
        Else
            sLog = sLog + Indent + "Installation of " + PackageName + " done!" + vbNewLine
        End If
    Else
        sLog = sLog + Indent + "Errors or warnings were raised while installing " + PackageName + ". Please check the log above." + vbNewLine
    End If

    sLog = sLog + Indent + "===================================" + vbNewLine
    
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & PackageName & VBA.Replace(VBA.Replace(VBA.Replace(VBA.Now(), ":", ""), "-", ""), " ", "") & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    If Simulate Then
        ThisApplication.Shell (sLogfile)
        If bOk Then
            If vbYes = Lime.MessageBox("Simulation of installation process completed. Please check the result in the recently opened logfile." & vbNewLine & vbNewLine & "Do you wish to proceed with the installation?", vbInformation + vbYesNo + vbDefaultButton2) Then
                Call lip.Install(PackageName, upgrade, False)
            End If
        Else
            Call Lime.MessageBox("Simulation of installation process completed. Errors occurred, please check the result in the recently opened logfile and take necessary actions before you try again.")
        End If
    Else
        If vbYes = Lime.MessageBox("Installation process completed. Remember to publish actionpads if needed. Do you want to open the logfile for the installation?", vbInformation + vbYesNo + vbDefaultButton1) Then
            ThisApplication.Shell (sLogfile)
        Else
            Debug.Print ("Logfile is available here: " & sLogfile)
        End If
    End If
    
    sLog = ""
    
    Application.MousePointer = 0

Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.Install")
End Sub

'Installs package from a zip-file. Input parameter: complete searchpath to the zip-file, including the filename
Public Sub InstallFromZip(ZipPath As String, Optional Simulate As Boolean = True)
On Error GoTo ErrorHandler
    
    Dim bOk As Boolean
    Dim sInstallPath As String
    
    Application.MousePointer = 11
    
    bOk = True
    sLog = ""
    IndentLenght = "  "

    'Check if valid path
    If VBA.Right(ZipPath, 4) = ".zip" Then
        If VBA.Dir(ZipPath) <> "" Then
            'Check if first use ever
            If Dir(WebFolder + "packages.json") = "" Then
                sLog = sLog + Indent + "No packages.json found, assuming fresh install" + vbNewLine
                Call InstallLIP
            End If

'           Copy file to actionpads\apps
            Dim PackageName As String
            Dim strArray() As String
            strArray = VBA.Split(ZipPath, "\")
            PackageName = VBA.Split(strArray(UBound(strArray)), ".")(0)
            sLog = sLog + Indent + "====== LIP Install: " + PackageName + " ======" + vbNewLine
            sLog = sLog + Indent + "Copying and unzipping file" + vbNewLine
            
            'TODO If prefix = app_ then change installpath to /apps else /packages
            If Left(PackageName, 4) = "app_" Then
                sInstallPath = Application.WebFolder & "apps\"
            Else
                sInstallPath = Application.WebFolder & DefaultInstallPath
            End If
            
            'Copy zip-file to the apps-folder if it's not already there
            'LJE Refactor
            'If ZipPath <> ThisApplication.WebFolder & "apps\" & PackageName & ".zip" Then
            If ZipPath <> sInstallPath & PackageName & ".zip" Then
                'LJE Refactor
                'Call VBA.FileCopy(ZipPath, ThisApplication.WebFolder & DefaultInstallPath & PackageName & ".zip")
                Call VBA.FileCopy(ZipPath, sInstallPath & PackageName & ".zip")
            End If
            
            
'           Unzip file
            'Refactor
            'Call Unzip(PackageName, ThisApplication.WebFolder & DefaultInstallPath)
            Call Unzip(PackageName, sInstallPath)

            'Get package information from json-file
            Dim Package As Object
            Dim sJSON As String
            Dim sLine As String
    
            'Look for package.json or app.json
            If VBA.Dir(sInstallPath & PackageName & "\" & "package.json") <> "" Then
                Open sInstallPath & PackageName & "\" & "package.json" For Input As #1
                
            ElseIf VBA.Dir(sInstallPath & PackageName & "\" & "app.json") <> "" Then
                Open sInstallPath & PackageName & "\" & "app.json" For Input As #1
            Else
                sLog = sLog + Indent + "Installation failed: couldn't find any package.json or app.json in the zip-file" + vbNewLine
                Call Application.MessageBox("ERROR: Installation failed: couldn't find any package.json or app.json in the zip-file")
                Application.Shell SaveLogFile(PackageName)
                Exit Sub
            End If

            Do Until EOF(1)
                Line Input #1, sLine
                sJSON = sJSON & sLine
            Loop

            Close #1

            Set Package = JSON.parse(sJSON)
            
            
            If Package.Exists("installPath") Then
                sInstallPath = ThisApplication.WebFolder & Package.Item("installPath") & "\"
            'LJE sätts högre upp
            'Else
            '    InstallPath = ThisApplication.WebFolder & DefaultInstallPath
            End If


            'Install dependencies
            If Package.Exists("dependencies") Then
                IncreaseIndent
                Call InstallDependencies(Package, Simulate)
                DecreaseIndent
            End If

            'LJE Refactor
            'If InstallPackageComponents(PackageName, 1, Package, InstallPath, Simulate) = False Then
            If InstallPackageComponents(PackageName, 1, Package, sInstallPath, Simulate) = False Then
                bOk = False
            End If
            
            If bOk Then
                If Simulate Then
                    sLog = sLog + Indent + "Simulation of " + PackageName + " done!" + vbNewLine
                Else
                    sLog = sLog + Indent + "Installation of " + PackageName + " done!" + vbNewLine
                End If
            Else
                sLog = sLog + Indent + "Errors or warnings were raised while installing " + PackageName + ". Please check the log above." + vbNewLine
                
            End If

            sLog = sLog + Indent + "===================================" + vbNewLine
            
            Dim sLogfile As String
            sLogfile = Application.TemporaryFolder & "\" & PackageName & VBA.Replace(VBA.Replace(VBA.Replace(VBA.Now(), ":", ""), "-", ""), " ", "") & ".txt"
            Open sLogfile For Output As #1
            Print #1, sLog
            Close #1
            
            If Simulate Then
                ThisApplication.Shell (sLogfile)
                If bOk Then
                    If vbYes = Lime.MessageBox("Simulation of installation process completed. Please check the result in the recently opened logfile." & vbNewLine & vbNewLine & "Do you wish to proceed with the installation?", vbInformation + vbYesNo + vbDefaultButton2) Then
                        Call lip.InstallFromZip(ZipPath, False)
                    End If
                Else
                    Call Lime.MessageBox("Simulation of installation process completed. Errors occurred, please check the result in the recently opened logfile and take necessary actions before you try again.")
                End If
            Else
                
                If vbYes = Lime.MessageBox("Installation process completed. Do you want to open the logfile for the installation?", vbInformation + vbYesNo + vbDefaultButton1) Then
                    ThisApplication.Shell (sLogfile)
                Else
                    Debug.Print ("Logfile is available here: " & sLogfile)
                End If
            End If
        Else
            Call Lime.MessageBox("Couldn't find file.")
        End If
    Else
        Call Lime.MessageBox("Path must end with .zip")
    End If
    
    sLog = ""
    
    Application.MousePointer = 0

Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallFromZip")
End Sub

Private Function SaveLogFile(strPackageName As String) As String
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & strPackageName & VBA.Replace(VBA.Replace(VBA.Replace(VBA.Now(), ":", ""), "-", ""), " ", "") & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    SaveLogFile = sLogfile
End Function

'Installs all packages defined in the packages.json file
Public Sub InstallFromPackageFile()
On Error GoTo ErrorHandler
    Dim LocalPackages As Object
    Dim LocalPackageName As Variant

    sLog = sLog + Indent + "Installing dependecies from packages.json file..." + vbNewLine
    Set LocalPackages = ReadPackageFile().Item("dependencies")
    If LocalPackages Is Nothing Then
        Exit Sub
    End If
    For Each LocalPackageName In LocalPackages.keys
        Call Install(CStr(LocalPackageName), True)
    Next LocalPackageName
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallFromPackageFile")
End Sub


Private Function InstallPackageComponents(PackageName As String, PackageVersion As Double, Package, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    
    Dim bOk As Boolean
    bOk = True

    'Install localizations
    If Package.Item("install").Exists("localize") = True Then
        sLog = sLog + Indent + "Adding localizations..." + vbNewLine
        IncreaseIndent
        If InstallLocalize(Package.Item("install").Item("localize"), Simulate) = False Then
            bOk = False
        End If
        DecreaseIndent

    End If

    'Install VBA
    If Package.Item("install").Exists("vba") = True Then
        sLog = sLog + Indent + "Adding VBA modules, forms and classes..." + vbNewLine
        IncreaseIndent
        If InstallVBAComponents(PackageName, Package.Item("install").Item("vba"), InstallPath, Simulate) = False Then
            bOk = False
        End If
        DecreaseIndent
    End If
    
    Dim sCreatedTables As String
    Dim sCreatedFields As String
    
    sCreatedTables = ""
    sCreatedFields = ""

    If Package.Item("install").Exists("tables") = True Then
        IncreaseIndent
        If InstallFieldsAndTables(Package.Item("install").Item("tables"), sCreatedTables, sCreatedFields) = False Then
            bOk = False
        End If
        DecreaseIndent
    End If
    
    If Package.Item("install").Exists("relations") = True Then
        IncreaseIndent
        If InstallRelations(Package.Item("install").Item("relations")) = False Then
            bOk = False
        End If
        DecreaseIndent
    End If
    
    If Simulate Then
        Call RollbackFieldsAndTables(sCreatedTables, sCreatedFields)
    End If

'    If Package.Item("install").Exists("sql") = True Then
'        IncreaseIndent
'        If InstallSQL(Package.Item("install").Item("sql"), PackageName, InstallPath, Simulate) = False Then
'            bOk = False
'        End If
'        DecreaseIndent
'    End If

    If Package.Item("install").Exists("files") = True Then
        IncreaseIndent
        If InstallFiles(Package.Item("install").Item("files"), PackageName, InstallPath, Simulate) = False Then
            bOk = False
        End If
        DecreaseIndent
    End If
    'Update packages.json
    If WriteToPackageFile(PackageName, CStr(PackageVersion), Simulate) = False Then
        bOk = False
    End If
    
    InstallPackageComponents = bOk
    
Exit Function
ErrorHandler:
    InstallPackageComponents = False
    Call UI.ShowError("lip.InstallPackageComponents")
End Function

Private Sub InstallDependencies(Package As Object, Simulate As Boolean)
On Error GoTo ErrorHandler
    Dim DependencyName As Variant
    Dim LocalPackage As Object
    sLog = sLog + Indent + "Dependencies found! Installing..." + vbNewLine
    IncreaseIndent
    For Each DependencyName In Package.Item("dependencies").keys()
        Set LocalPackage = FindPackageLocally(CStr(DependencyName))
        If LocalPackage Is Nothing Then
            sLog = sLog + Indent + "Installing dependency: " + CStr(DependencyName) + vbNewLine
            Call Install(CStr(DependencyName), Simulate)
        ElseIf CDbl(VBA.Replace(LocalPackage.Item(DependencyName), ".", ",")) < CDbl(VBA.Replace(Package.Item("dependencies").Item(DependencyName), ".", ",")) Then
            Call Install(CStr(DependencyName), True, Simulate)
        Else
        End If
    Next DependencyName
    Call DecreaseIndent
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallDependencies")
End Sub


Private Function SearchForPackageInStores(PackageName As String) As Object
On Error GoTo ErrorHandler
        
    Set SearchForPackageInStores = SearchForPackageInOnlineStores(PackageName)
    
    If SearchForPackageInStores Is Nothing Then
        Set SearchForPackageInStores = SearchForPackageInLocalStores(PackageName)
        If SearchForPackageInStores Is Nothing Then
            'If we've reached this code, package wasn't found
            Debug.Print Indent + ("Package/App '" & PackageName & "' not found!")
            Set SearchForPackageInStores = Nothing
        End If
    End If

Exit Function
ErrorHandler:
    Set SearchForPackageInStores = Nothing
    Call UI.ShowError("lip.SearchForPackageInStores")
End Function

'LJE Search for package in online stores
Public Function SearchForPackageInOnlineStores(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim sJSON As String
    Dim oJSON As Object
    Dim oStores As Object
    Dim Path As String
    Dim oStore As Variant
    'LJE changed to onlinestores
    'Set oPackages = ReadPackageFile.Item("stores")
    Set oStores = ReadPackageFile.Item("onlinestores")

    'Loop through packagestores from packages.json
    For Each oStore In oStores


        Path = oStores.Item(oStore)
        sLog = sLog + Indent + ("Looking for package at store '" & oStore & "'") + vbNewLine
        
        sJSON = getJSON(Path + PackageName + "/")

        If sJSON <> "" Then
            sJSON = VBA.Left(sJSON, VBA.Len(sJSON) - 1) & ",""source"":""" & oStores.Item(oStore) & """}" 'Add a source node so we know where the package exists
        End If

        Set oJSON = ParseJson(sJSON) 'Create a JSON object from the string

        If Not oJSON Is Nothing Then
            If oJSON.Item("error") = "" Then
                'Package found, make sure the install node exists
                If Not oJSON.Item("install") Is Nothing Then
                    sLog = sLog + Indent + ("Package '" & PackageName & "' found on store '" & oStore & "'") + vbNewLine
                    Set SearchForPackageInOnlineStores = oJSON
                    Exit Function
                Else
                    sLog = sLog + Indent + ("Package '" & PackageName & "' found on store '" & oStore & "' but has no valid install instructions!") + vbNewLine
                    Set SearchForPackageInOnlineStores = oJSON
                    Exit Function
                End If
            End If
        End If
    Next
    
    'If we've reached this code, package wasn't found
    sLog = sLog + Indent + ("Package '" & PackageName & "' not found!") + vbNewLine
    Set SearchForPackageInOnlineStores = Nothing

Exit Function
ErrorHandler:
    Set SearchForPackageInOnlineStores = Nothing
    Call UI.ShowError("lip.SearchForPackageInOnlineStores")
End Function


'LJE Search for package in local stores
'Should be a local path where folders are named after packages
'LJE TEST
Public Function SearchForPackageInLocalStores(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim oStores As Object
    Dim oStore As Variant
    Dim Path As String
    Dim FileSystem As Object
    Dim oJSON As Object
    
    Set oStores = ReadPackageFile.Item("localstores")
    'TODO Test if the oStores is ok
    'TODO Test with multiple local stores
    
    'Loop through localstores from packages.json
    For Each oStore In oStores
        
        Path = oStores.Item(oStore)
        Debug.Print Indent + ("Looking for '" & PackageName & "' at store '" & oStore & "'")
        
        Dim FileSystemObj As Object
        Dim startFolder As Object
        Dim fld As Object
        
        Set FileSystemObj = CreateObject("Scripting.FileSystemObject")
        'LJE backslash needs to be handled - see trello item.
        'LJE TODO Check if store path is ok
        Set startFolder = FileSystemObj.GetFolder(Path)
        
        
        For Each fld In startFolder.SubFolders
            If LCase(fld.Name) = LCase(PackageName) Then
                Dim sJSON As String
                Dim sLine As String
                
                Open fld.Path & "\" & "app.json" For Input As #1
                        
                Do Until EOF(1)
                    Line Input #1, sLine
                    sJSON = sJSON & sLine
                Loop
                
                
                If sJSON <> "" Then
                    Dim sPathToLocalPackage As String
                    sPathToLocalPackage = VBA.Replace(fld.Path, "\", "\\")
                    sJSON = VBA.Left(sJSON, VBA.Len(sJSON) - 1) & ",""localsource"":""" & sPathToLocalPackage + "\" + fld.Name + """}"   'Add a source node so we know where the package exists
                End If
    
                Close #1
                
                Set oJSON = ParseJson(sJSON) 'Create a JSON object from the string
                
                If Not oJSON.Item("install") Is Nothing Then
                    Debug.Print Indent + ("Package/App '" & PackageName & "' found in local store '" & oStore & "'")
                    Set SearchForPackageInLocalStores = oJSON
                    Exit Function
                Else
                    Debug.Print Indent + ("Package/App '" & PackageName & "' found in local store '" & oStore & "' but has no valid install instructions!")
                    Set SearchForPackageInLocalStores = Nothing
                    Exit Function
                End If
                
            End If
        Next
    Next
       
    'If we've reached this code, package wasn't found
    Debug.Print Indent + ("Package/App '" & PackageName & "' not found in local stores!")
    Set SearchForPackageInLocalStores = Nothing
    
    Exit Function
ErrorHandler:
    Set SearchForPackageInLocalStores = Nothing
    Call UI.ShowError("lip.SearchForPackageInLocalStores")

End Function

Private Function CheckForLocalInstalledPackage(PackageName As String, PackageVersion As Double) As Boolean
On Error GoTo ErrorHandler
    Dim LocalPackages As Object
    Dim LocalPackage As Object
    Dim LocalPackageVersion As Double
    Dim LocalPackageName As Variant

    Set LocalPackage = FindPackageLocally(PackageName)

    If Not LocalPackage Is Nothing Then
        LocalPackageVersion = CDbl(VBA.Replace(LocalPackage.Item(PackageName), ".", ","))
        If PackageVersion = LocalPackageVersion Then
            sLog = sLog + Indent + "Current version of" + PackageName + " is already installed, please use the upgrade command to reinstall package" + vbNewLine
            sLog = sLog + Indent + "===================================" + vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        ElseIf PackageVersion > LocalPackageVersion Then
            sLog = sLog + Indent + "Package " + PackageName + " is already installed, please use the upgrade command to upgrade package from " + Format(LocalPackageVersion, "0.0") + " -> " + Format(PackageVersion, "0.0") + vbNewLine
            sLog = sLog + Indent + "===================================" + vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        Else
            sLog = sLog + Indent + "A newer version of " + PackageName + " is already installed. Remote: " + Format(PackageVersion, "0.0") + " ,Local: " + Format(LocalPackageVersion, "0.0") + ". Please use the upgrade command to reinstall package" + vbNewLine
            sLog = sLog + Indent + "===================================" + vbNewLine
            CheckForLocalInstalledPackage = True
            Exit Function
        End If
    End If
    CheckForLocalInstalledPackage = False
Exit Function
ErrorHandler:
    Call UI.ShowError("lip.CheckForLocalInstalledPackages")
End Function

Private Function getJSON(sURL As String) As String
On Error GoTo ErrorHandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim oXHTTP As Object
    Dim s As String
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    oXHTTP.Open "GET", sURL + "?" + qs, False
    oXHTTP.Send
    getJSON = oXHTTP.responseText
Exit Function
ErrorHandler:
    getJSON = ""
End Function

Private Function ParseJson(sJSON As String) As Object
On Error GoTo ErrorHandler
    Dim oJSON As Object
    Set oJSON = JSON.parse(sJSON)
    Set ParseJson = oJSON
Exit Function
ErrorHandler:
    Set ParseJson = Nothing
    Call UI.ShowError("lip.parseJSON")
End Function

Private Function findNewestVersion(oVersions As Object) As Double
On Error GoTo ErrorHandler
    Dim NewestVersion As Double
    Dim Version As Variant
    NewestVersion = -1

    For Each Version In oVersions
        If CDbl(VBA.Replace(Version.Item("version"), ".", ",")) > NewestVersion Then
            NewestVersion = CDbl(VBA.Replace(Version.Item("version"), ".", ","))
        End If
    Next Version
    findNewestVersion = NewestVersion
Exit Function
ErrorHandler:
    findNewestVersion = -1
    Call UI.ShowError("lip.findNewestVersion")
End Function

Private Function InstallLocalize(oJSON As Object, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim Localize As Variant
    bOk = True
    
    For Each Localize In oJSON
        If AddOrCheckLocalize( _
            Localize.Item("owner"), _
            Localize.Item("context"), _
            "", _
            Localize.Item("en_us"), _
            Localize.Item("sv"), _
            Localize.Item("no"), _
            Localize.Item("fi"), _
            Simulate _
        ) = False Then
            bOk = False
        End If
    Next Localize
    
    InstallLocalize = bOk
    
Exit Function
ErrorHandler:
    InstallLocalize = False
    Call UI.ShowError("lip.InstallLocalize")
End Function

Private Function InstallFiles(oJSON As Object, PackageName As String, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim FSO As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim File As Variant
    
    bOk = True

    For Each File In oJSON
        FromPath = InstallPath & PackageName & "\" & File
        ToPath = WebFolder & File

        If Right(FromPath, 1) = "\" Then
            FromPath = Left(FromPath, Len(FromPath) - 1)
        End If
        If Right(ToPath, 1) = "\" Then
            ToPath = Left(ToPath, Len(ToPath) - 1)
        End If
        Set FSO = CreateObject("scripting.filesystemobject")

        FSO.CopyFolder Source:=FromPath, Destination:=ToPath
        On Error Resume Next 'It is a beautiful languge
        If Simulate Then
            VBA.Kill ToPath
        Else
            VBA.Kill FromPath
        End If
        On Error GoTo ErrorHandler
    Next File
    
    InstallFiles = bOk

ErrorHandler:
    InstallFiles = False
    Call UI.ShowError("lip.InstallFiles")
End Function

'Private Function InstallSQL(oJSON As Object, PackageName As String, InstallPath As String) As Boolean
'On Error GoTo ErrorHandler
'    Dim bOk As Boolean
'    Dim SQL As Variant
'    Dim Path As String
'    Dim RelPath As String
'
'    bOk = True
'
'    slog=slog+ Indent + "Installing SQL..." +vbnewline
'    IncreaseIndent
'    For Each SQL In oJSON
'        RelPath = Replace(SQL.Item("relPath"), "/", "\")
'        Path = InstallPath & PackageName & "\" & RelPath
'        If CreateSQLProcedure(Path, SQL.Item("name"), SQL.Item("type")) = False Then
'            bOk = False
'        End If
'    Next SQL
'    DecreaseIndent
'    InstallSQL = bOk
'Exit Function
'ErrorHandler:
'    InstallSQL = False
'    Call UI.ShowError("lip.InstallSQL")
'End Function
'
'Private Function CreateSQLProcedure(Path As String, Name As String, ProcType As String) As Boolean
'    Dim bOk As Boolean
'    Dim oProc As New LDE.Procedure
'    Dim strSQL As String
'    Dim sLine As String
'    Dim sErrormessage As String
'
'    bOk = True
'    strSQL = ""
'    sErrormessage = ""
'
'    Open Path For Input As #1
'        Do Until EOF(1)
'            Line Input #1, sLine
'            strSQL = strSQL & sLine & vbNewLine
'        Loop
'        Close #1
'
'        Set oProc = Database.Procedures("csp_lip_installSQL")
'        If Not oProc Is Nothing Then
'            oProc.Parameters("@@sql") = strSQL
'            oProc.Parameters("@@name") = Name
'            oProc.Parameters("@@type") = ProcType
'            oProc.Execute (False)
'
'            sErrormessage = oProc.Parameters("@@errormessage").OutputValue
'
'            If sErrormessage <> "" Then
'                slog=slog+ Indent + (sErrormessage)+vbnewline
'                bOk = False
'            Else
'                slog=slog+ Indent + ("'" & Name & "'" & " added.")+vbnewline
'            End If
'
'        Else
'            bOk = False
'            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_installSQL'. Please make sure this procedure exists in the database and restart LDC.")
'        End If
'
'        CreateSQLProcedure = bOk
'
'Exit Function
'ErrorHandler:
'    CreateSQLProcedure = False
'    Call UI.ShowError("lip.CreateSQLProcedure")
'End Function

Private Function InstallFieldsAndTables(oJSON As Object, ByRef sCreatedTables As String, ByRef sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim table As Object
    Dim oProc As LDE.Procedure
    Dim field As Object
    Dim idtable As Long
    Dim iddescriptiveexpression As Long
    Dim oItem As Variant

    Dim localname_singular As String
    Dim localname_plural As String
    Dim ErrorMessage As String
    Dim warningmessage As String
    
    bOk = True

    sLog = sLog + Indent + "Adding fields and tables..." + vbNewLine
    IncreaseIndent

    For Each table In oJSON
        localname_singular = ""
        localname_plural = ""
        ErrorMessage = ""
        idtable = -1

        Set oProc = Database.Procedures("csp_lip_createtable")
        oProc.Timeout = 299

        If Not oProc Is Nothing Then

            sLog = sLog + Indent + "Add table: " + table.Item("name") + vbNewLine

            oProc.Parameters("@@tablename").InputValue = table.Item("name")

            'Add localnames singular
            If table.Exists("localname_singular") Then
                For Each oItem In table.Item("localname_singular")
                    If oItem <> "" Then
                        localname_singular = localname_singular + VBA.Trim(oItem) + ":" + VBA.Trim(table.Item("localname_singular").Item(oItem)) + ";"
                    End If
                Next
                oProc.Parameters("@@localname_singular").InputValue = localname_singular
            End If

            'Add localnames plural
            If table.Exists("localname_plural") Then
                For Each oItem In table.Item("localname_plural")
                    If oItem <> "" Then
                        localname_plural = localname_plural + VBA.Trim(oItem) + ":" + VBA.Trim(table.Item("localname_plural").Item(oItem)) + ";"
                    End If
                Next
                oProc.Parameters("@@localname_plural").InputValue = localname_plural
            End If

            Call oProc.Execute(False)

            ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
            warningmessage = oProc.Parameters("@@warningMessage").OutputValue

            idtable = oProc.Parameters("@@idtable").OutputValue
            iddescriptiveexpression = oProc.Parameters("@@iddescriptiveexpression").OutputValue
            
            If idtable <> -1 Then
                sCreatedTables = sCreatedTables + CStr(idtable) + ";"
            End If

            If warningmessage <> "" Then
                IncreaseIndent
                sLog = sLog + Indent + (warningmessage) + vbNewLine
                DecreaseIndent
            End If
            
            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                IncreaseIndent
                sLog = sLog + Indent + (ErrorMessage) + vbNewLine
                bOk = False
                DecreaseIndent
            Else
                sLog = sLog + Indent + ("Table """ & table.Item("name") & """ installed.") + vbNewLine
            End If

            ' Create fields
            IncreaseIndent
            If table.Exists("fields") Then
                For Each field In table.Item("fields")
                    sLog = sLog + Indent + "Add field: " + field.Item("name") + vbNewLine
                    If AddField(table.Item("name"), field, sCreatedFields) = False Then
                        bOk = False
                    End If
                Next field
            End If

            'Set table attributes(must be done AFTER fields has been created in order to be able to set descriptive expression)
            'Only set attributes if table was created
            If idtable <> -1 Then
                If SetTableAttributes(table, idtable, iddescriptiveexpression) = False Then
                    bOk = False
                End If
            End If

            DecreaseIndent

        Else
            bOk = False
            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_createtable'. Please make sure this procedure exists in the database and restart LDC.")
        End If

    Next table
    DecreaseIndent

    Set oProc = Nothing
    
    InstallFieldsAndTables = bOk

    Exit Function
ErrorHandler:
    Set oProc = Nothing
    InstallFieldsAndTables = False
    Call UI.ShowError("lip.InstallFieldsAndTables")
End Function


Private Function AddField(tableName As String, field As Object, ByRef sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim oProc As New LDE.Procedure
    Dim ErrorMessage As String
    Dim warningmessage As String
    Dim fieldLocalnames As String
    Dim separatorLocalnames As String
    Dim limevalidationtextLocalnames As String
    Dim commentLocalnames As String
    Dim tooltipLocalnames As String
    Dim oItem As Variant
    Dim optionItems As Variant
    Dim idfield As Long
    
    bOk = True
    ErrorMessage = ""
    warningmessage = ""
    fieldLocalnames = ""
    separatorLocalnames = ""
    limevalidationtextLocalnames = ""
    commentLocalnames = ""
    tooltipLocalnames = ""
    idfield = -1
    
    Set oProc = Database.Procedures("csp_lip_createfield")
    oProc.Timeout = 299

    If Not oProc Is Nothing Then
        oProc.Parameters("@@tablename").InputValue = tableName
        oProc.Parameters("@@fieldname").InputValue = field.Item("name")

        'Add localnames
        If field.Exists("localname") Then
            For Each oItem In field.Item("localname")
                If oItem <> "" Then
                    fieldLocalnames = fieldLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("localname").Item(oItem)) + ";"
                End If
            Next
            oProc.Parameters("@@localname").InputValue = fieldLocalnames
        End If

        'Add attributes
        If field.Exists("attributes") Then
            For Each oItem In field.Item("attributes")
                If oItem <> "" Then
                    If Not oProc.Parameters.Lookup("@@" & oItem, lkLookupProcedureParameterByName) Is Nothing Then
                        oProc.Parameters("@@" & oItem).InputValue = field.Item("attributes").Item(oItem)
                    Else
                        sLog = sLog + Indent + ("No support for setting field attribute " & oItem) + vbNewLine
                    End If
                End If
            Next
        End If

        'Add separator
        If field.Exists("separator") Then
            For Each oItem In field.Item("separator")
                separatorLocalnames = separatorLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("separator").Item(oItem)) + ";"
            Next
            oProc.Parameters("@@separator").InputValue = separatorLocalnames
        End If
        
        'Add limevalidationtext
        If field.Exists("limevalidationtext") Then
            For Each oItem In field.Item("limevalidationtext")
                limevalidationtextLocalnames = limevalidationtextLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("limevalidationtext").Item(oItem)) + ";"
            Next
            oProc.Parameters("@@limevalidationtext").InputValue = limevalidationtextLocalnames
        End If
        
        'Add comment
        If field.Exists("comment") Then
            For Each oItem In field.Item("comment")
                commentLocalnames = commentLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("comment").Item(oItem)) + ";"
            Next
            oProc.Parameters("@@comment").InputValue = commentLocalnames
        End If
        
        'Add tooltip (description)
        If field.Exists("description") Then
            For Each oItem In field.Item("description")
                tooltipLocalnames = tooltipLocalnames + VBA.Trim(oItem) + ":" + VBA.Trim(field.Item("description").Item(oItem)) + ";"
            Next
            oProc.Parameters("@@description").InputValue = tooltipLocalnames
        End If

        Dim strOptions As String
        strOptions = ""
        'Add options
        If field.Exists("options") Then
            For Each optionItems In field.Item("options")
                strOptions = strOptions + "["
                For Each oItem In optionItems
                    strOptions = strOptions + VBA.Trim(oItem) + ":" + VBA.Trim(optionItems.Item(oItem)) + ";"
                Next
                strOptions = strOptions + "]"
            Next
            oProc.Parameters("@@optionlist").InputValue = strOptions
        End If

        Call oProc.Execute(False)
        ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
        warningmessage = oProc.Parameters("@@warningMessage").OutputValue
        
        idfield = oProc.Parameters("@@idfield").OutputValue
        
        If idfield <> -1 Then
            sCreatedFields = sCreatedFields + CStr(idfield) + ";"
        End If
        
        'Log warnings
        If warningmessage <> "" Then
            IncreaseIndent
            sLog = sLog + Indent + (warningmessage) + vbNewLine
            DecreaseIndent
        End If
        
        'If errormessage is set, something went wrong
        If ErrorMessage <> "" Then
            IncreaseIndent
            sLog = sLog + Indent + (ErrorMessage) + vbNewLine
            DecreaseIndent
            bOk = False
        Else
            sLog = sLog + Indent + ("Field """ & field.Item("name") & """ installed.") + vbNewLine
        End If
    Else
        bOk = False
        Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_createfield'. Please make sure this procedure exists in the database and restart LDC.")
    End If
    Set oProc = Nothing
    AddField = bOk

    Exit Function
ErrorHandler:
    Set oProc = Nothing
    AddField = False
    Call UI.ShowError("lip.AddField")
End Function

Private Function SetTableAttributes(ByRef table As Object, idtable As Long, iddescriptiveexpression As Long) As Boolean
On Error GoTo ErrorHandler

    Dim bOk As Boolean
    Dim oProcAttributes As LDE.Procedure
    Dim oItem As Variant
    Dim ErrorMessage As String
    Dim warningmessage As String
    
    bOk = True
    ErrorMessage = ""
    warningmessage = ""

    If table.Exists("attributes") Then

        Set oProcAttributes = Application.Database.Procedures("csp_lip_settableattributes")
        oProcAttributes.Timeout = 299

        If Not oProcAttributes Is Nothing Then

            sLog = sLog + Indent + "Adding attributes for table: " + table.Item("name") + vbNewLine

            oProcAttributes.Parameters("@@tablename").InputValue = table.Item("name")
            oProcAttributes.Parameters("@@idtable").InputValue = idtable
            oProcAttributes.Parameters("@@iddescriptiveexpression").InputValue = iddescriptiveexpression

            For Each oItem In table.Item("attributes")
                If oItem <> "" Then
                    If Not oProcAttributes.Parameters.Lookup("@@" & oItem, lkLookupProcedureParameterByName) Is Nothing Then
                        oProcAttributes.Parameters("@@" & oItem).InputValue = table.Item("attributes").Item(oItem)
                    Else
                        sLog = sLog + Indent + ("No support for setting table attribute " & oItem) + vbNewLine
                    End If
                End If
            Next

            Call oProcAttributes.Execute(False)

            ErrorMessage = oProcAttributes.Parameters("@@errorMessage").OutputValue
            warningmessage = oProcAttributes.Parameters("@@warningMessage").OutputValue
            
            If warningmessage <> "" Then
                sLog = sLog + Indent + (warningmessage) + vbNewLine
            End If

            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                sLog = sLog + Indent + (ErrorMessage) + vbNewLine
                bOk = False
            Else
                sLog = sLog + Indent + ("Attributes for table """ & table.Item("name") & """ set.") + vbNewLine
            End If

        Else
            bOk = False
            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_settableattributes'. Please make sure this procedure exists in the database and restart LDC.")
        End If
    End If

    Set oProcAttributes = Nothing
    
    SetTableAttributes = bOk

    Exit Function
ErrorHandler:
    Set oProcAttributes = Nothing
    SetTableAttributes = False
    Call UI.ShowError("lip.SetTableAttributes")
End Function

Private Function DownloadFile(PackageName As String, Path As String, InstallPath As String) As String
On Error GoTo ErrorHandler
    Dim qs As String
    qs = CStr(Rnd() * 1000000#)
    Dim downloadURL As String
    Dim myURL As String
    Dim oStream As Object
    downloadURL = Path + PackageName + "/download/"

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", downloadURL + "?" + qs, False
    WinHttpReq.Send

    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile InstallPath + PackageName + ".zip", 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
    DownloadFile = ""
    Exit Function
ErrorHandler:
    DownloadFile = "Couldn't download file from " & downloadURL & vbCrLf & vbCrLf & Err.Description
End Function

Private Sub Unzip(PackageName As String, InstallPath As String)
On Error GoTo ErrorHandler
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = InstallPath + PackageName + ".zip"
    FileNameFolder = InstallPath & PackageName & "\"

    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    'Delete files
    FSO.DeleteFile FileNameFolder & "*.*", True
    'Delete subfolders
    FSO.DeleteFolder FileNameFolder & "*.*", True

    'Make the normal folder in DefPath
    MkDir FileNameFolder

    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items

    'Delete zip-file
    FSO.DeleteFile Fname, True

    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.Unzip")
End Sub

Private Function InstallVBAComponents(PackageName As String, VBAModules As Object, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    bOk = True
    Dim VBAModule As Variant
    For Each VBAModule In VBAModules
        If addModule(PackageName, VBAModule.Item("name"), VBAModule.Item("relPath"), InstallPath, Simulate) = False Then
            bOk = False
        Else
            Debug.Print Indent + "Added " + VBAModule.Item("name")
        End If
    Next VBAModule
    InstallVBAComponents = bOk
    Exit Function
ErrorHandler:
    InstallVBAComponents = False
    Call UI.ShowError("lip.InstallVBAComponents")
End Function

Private Function addModule(PackageName As String, ModuleName As String, RelPath As String, InstallPath As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    bOk = True
    If PackageName <> "" And ModuleName <> "" Then
        Dim VBComps As Object
        Dim Path As String
        Dim tempModuleName As String

        Set VBComps = Application.VBE.ActiveVBProject.VBComponents

        Path = InstallPath + PackageName + "\" + Replace(RelPath, "/", "\")
        
        If VBA.Dir(Path) <> "" Then
            If ComponentExists(ModuleName, VBComps) = True Then
                If vbYes = Lime.MessageBox("Do you want to replace existing VBA-module """ & ModuleName & """?", vbYesNo + vbDefaultButton2 + vbQuestion) Then
                    tempModuleName = LCO.GenerateGUID
                    tempModuleName = VBA.Replace(VBA.Mid(tempModuleName, 2, VBA.Len(tempModuleName) - 2), "-", "")
                    tempModuleName = VBA.Left("OLD_" & tempModuleName, 30)
                    
                    If Not Simulate Then
                        VBComps.Item(ModuleName).Name = tempModuleName
                    End If
                    
                    If vbYes = Lime.MessageBox("Do you want to delete the old module?", vbYesNo + vbDefaultButton2 + vbQuestion) Then
                        If Not Simulate Then
                            Call VBComps.Remove(VBComps.Item(tempModuleName))
                        End If
                    Else
                        Call Lime.MessageBox("Old module is saved with the name """ & tempModuleName & """", vbInformation)
                        sLog = sLog + Indent + ("Old module is saved with the name """ & tempModuleName & """") + vbNewLine
                    End If
                    
                    If Not Simulate Then
                        Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
                    End If
                    sLog = sLog + Indent + "VBA added: " + ModuleName + vbNewLine
                Else
                    sLog = sLog + Indent + ("Module """ & ModuleName & """ already exists and have not been replaced.") + vbNewLine
                End If
            Else
                
                If Not Simulate Then
                    Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
                End If
                sLog = sLog + Indent + "Added " + ModuleName + vbNewLine
            End If
        Else
            sLog = sLog + Indent + "Module """ & ModuleName & """ can't be added. File does not exists." + vbNewLine
        End If
        
    Else
        bOk = False
        sLog = sLog + (Indent + "Detected invalid package- or modulename while installing """ + RelPath + """") + vbNewLine
    End If
    addModule = bOk
    Exit Function
ErrorHandler:
    addModule = False
    Call UI.ShowError("lip.addModule")

    sLog = sLog + Indent + "Couldn't add module " + ModuleName + vbNewLine
End Function

Private Function ComponentExists(ComponentName As String, VBComps As Object) As Boolean
On Error GoTo ErrorHandler
    Dim VBComp As Variant

    For Each VBComp In VBComps
        If VBComp.Name = ComponentName Then
             ComponentExists = True
             Exit Function
        End If
    Next VBComp

    ComponentExists = False

    Exit Function
ErrorHandler:
    Call UI.ShowError("lip.ComponentExists")
End Function

Private Function WriteToPackageFile(PackageName As String, Version As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim oJSON As Object
    Dim fs As Object
    Dim a As Object
    Dim Line As Variant
    
    bOk = True
    Set oJSON = ReadPackageFile

    oJSON.Item("dependencies").Item(PackageName) = Version
    
    If Not Simulate Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(WebFolder + "packages.json", True)
        For Each Line In Split(PrettyPrintJSON(JSON.toString(oJSON)), vbCrLf)
            Line = VBA.Replace(Line, "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with packagestores URLs
            a.WriteLine Line
        Next Line
        a.Close
    End If
    
    WriteToPackageFile = bOk
    Exit Function
ErrorHandler:
    WriteToPackageFile = False
    Call UI.ShowError("lip.WriteToPackageFile")
End Function

Private Function PrettyPrintJSON(JSON As String) As String
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim Indent As String
    Dim PrettyJSON As String
    Dim InsideQuotation As Boolean

    For i = 1 To Len(JSON)
        Select Case Mid(JSON, i, 1)
            Case """"
                PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                If InsideQuotation = False Then
                    InsideQuotation = True
                Else
                    InsideQuotation = False
                End If
            Case "{", "["
                If InsideQuotation = False Then
                    Indent = Indent + "    " ' Add to indentation
                    PrettyJSON = PrettyJSON + "{" + vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case "}", "["
                If InsideQuotation = False Then
                    Indent = Left(Indent, Len(Indent) - 4) 'Remove indentation
                    PrettyJSON = PrettyJSON + vbCrLf + Indent + "}"
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case ","
                If InsideQuotation = False Then
                    PrettyJSON = PrettyJSON + "," + vbCrLf + Indent
                Else
                    PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
                End If
            Case Else
                PrettyJSON = PrettyJSON + Mid(JSON, i, 1)
        End Select
    Next i
    PrettyPrintJSON = PrettyJSON

    Exit Function
ErrorHandler:
    PrettyPrintJSON = ""
    Call UI.ShowError("lip.PrettyPrintJSON")
End Function

Private Function ReadPackageFile() As Object
On Error GoTo ErrorHandler
    Dim sJSON As String
    Dim oJSON As Object
    sJSON = getJSON(WebFolder + "packages.json")

    If sJSON = "" Then
        sLog = sLog + Indent + "Error: No packages.json found!" + vbNewLine
        Set ReadPackageFile = Nothing
        Exit Function
    End If

    Set oJSON = JSON.parse(sJSON)
    Set ReadPackageFile = oJSON

    Exit Function
ErrorHandler:
    Set ReadPackageFile = Nothing
    Call UI.ShowError("lip.ReadPackageFile")
End Function

Private Function FindPackageLocally(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim InstalledPackages As Object
    Dim Package As Object
    Dim ReturnDict As New Scripting.Dictionary
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile

    If Not oPackageFile Is Nothing Then

        If oPackageFile.Exists("dependencies") Then
            Set InstalledPackages = oPackageFile.Item("dependencies")
            If InstalledPackages.Exists(PackageName) = True Then
                Call ReturnDict.Add(PackageName, InstalledPackages.Item(PackageName))
                Set FindPackageLocally = ReturnDict
                Exit Function
            End If
        Else
            sLog = sLog + Indent + ("Couldn't find dependencies in packages.json") + vbNewLine
        End If

    End If

    Set FindPackageLocally = Nothing
    Exit Function
ErrorHandler:
    Set FindPackageLocally = Nothing
    Call UI.ShowError("lip.FindPackageLocally")
End Function
'LJE TODO Refactor with helper method to write json
'TEST
Public Sub CreateANewPackageFile()
On Error GoTo ErrorHandler
    Dim fs As Object
    Dim a As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WebFolder + "packages.json", True)
    a.WriteLine ("{")
    'LJE VersionHandling
    'TODO write to GitHub
    a.WriteLine ("    ""lipversion"":0.1,")
    'LJE Should perhaps have two different objects - one onlinestore and one localstore
    a.WriteLine ("    ""onlinestores"":{")
    a.WriteLine ("        ""PackageStore"":""http://api.lime-bootstrap.com/packages/"",")
    a.WriteLine ("        ""Bootstrap Appstore"":""http://api.lime-bootstrap.com/apps/""")
    a.WriteLine ("    },")
    a.WriteLine ("    ""localstores"":{")
    a.WriteLine ("    },")
    a.WriteLine ("    ""dependencies"":{")
    a.WriteLine ("    }")
    a.WriteLine ("}")
    a.Close
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.CreateNewPackageFile")
End Sub

Public Function GetAllInstalledPackages() As String
On Error GoTo ErrorHandler
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile()

    If Not oPackageFile Is Nothing Then
        GetAllInstalledPackages = JSON.toString(oPackageFile)
    Else
        GetAllInstalledPackages = "{}"
        sLog = sLog + Indent + "Couldn't find dependencies in packages.json" + vbNewLine
    End If

    Exit Function
ErrorHandler:
    Call UI.ShowError("lip.GetInstalledPackages")
End Function

Public Sub InstallLIP()
On Error GoTo ErrorHandler
    Dim InstallPath As String
    
    sLog = ""

    sLog = sLog + Indent + "Creating a new packages.json file..." + vbNewLine
    Call CreateANewPackageFile
    Dim FSO As New FileSystemObject
    InstallPath = ThisApplication.WebFolder & DefaultInstallPath
    If Not FSO.FolderExists(InstallPath) Then
        FSO.CreateFolder InstallPath
    End If

    sLog = sLog + Indent + "Installing JSON-lib..." + vbNewLine
    Dim strDownloadError
    strDownloadError = DownloadFile("vba_json", BaseURL + AppStoreApiURL, InstallPath)
    If strDownloadError = "" Then
        Call Unzip("vba_json", InstallPath)
    
        Call addModule("vba_json", "JSON", "JSON.bas", InstallPath, False)
        Call addModule("vba_json", "cStringBuilder", "cStringBuilder.cls", InstallPath, False)
    
        Call WriteToPackageFile("vba_json", "1", False)
    
        sLog = sLog + Indent + "Install of LIP complete!" + vbNewLine
    Else
        sLog = sLog + Indent + "Could not download the package vba_json from the Appstore: " + BaseURL + AppStoreApiURL
    End If
    Dim sLogfile As String
    sLogfile = Application.TemporaryFolder & "\" & "lip" & VBA.Replace(VBA.Replace(VBA.Replace(VBA.Now(), ":", ""), "-", ""), " ", "") & ".txt"
    Open sLogfile For Output As #1
    Print #1, sLog
    Close #1
    
    Application.Shell sLogfile
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallLIP")
End Sub

Private Function AddOrCheckLocalize(sOwner As String, sCode As String, sDescription As String, sEN_US As String, sSV As String, sNO As String, sFI As String, Simulate As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim oFilter As New LDE.Filter
    Dim oRecs As New LDE.Records

    Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    Call oFilter.AddCondition("code", lkOpEqual, sCode)
    Call oFilter.AddOperator(lkOpAnd)

    If oFilter.HitCount(Database.Classes("localize")) = 0 Then
        sLog = sLog + Indent + "Localization " & sOwner & "." & sCode & " not found, creating new!" + vbNewLine
        If Not Simulate Then
            Dim oRec As New LDE.Record
            Call oRec.Open(Database.Classes("localize"))
            oRec.Value("owner") = sOwner
            oRec.Value("code") = sCode
            oRec.Value("context") = sDescription
            oRec.Value("sv") = sSV
            oRec.Value("en_us") = sEN_US
            oRec.Value("no") = sNO
            oRec.Value("fi") = sFI
            Call oRec.Update
        End If
    ElseIf oFilter.HitCount(Database.Classes("localize")) = 1 Then
        sLog = sLog + Indent + "Updating localization " + sOwner + "." + sCode + vbNewLine
        
        If Not Simulate Then
            Call oRecs.Open(Database.Classes("localize"), oFilter)
            oRecs(1).Value("owner") = sOwner
            oRecs(1).Value("code") = sCode
            oRecs(1).Value("context") = sDescription
            oRecs(1).Value("sv") = sSV
            oRecs(1).Value("en_us") = sEN_US
            oRecs(1).Value("no") = sNO
            oRecs(1).Value("fi") = sFI
            Call oRecs.Update
        End If

    Else
        sLog = sLog + Indent + "There are multiple copies of " & sOwner & "." & sCode & ". Fix this and try again."
    End If

    Set Localize.dicLookup = Nothing
    AddOrCheckLocalize = True
    Exit Function
ErrorHandler:
    sLog = sLog + Indent + ("Error while validating or adding Localize") + vbNewLine
    AddOrCheckLocalize = False
End Function

Private Sub IncreaseIndent()
On Error GoTo ErrorHandler
    Indent = Indent + IndentLenght
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.IncreaseIndent")
End Sub

Private Sub DecreaseIndent()
On Error GoTo ErrorHandler

    If Len(Indent) - Len(IndentLenght) > 0 Then
        Indent = Left(Indent, Len(Indent) - Len(IndentLenght))
    Else
        Indent = ""
    End If
    
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.DecreaseIndent")
End Sub

Private Function InstallRelations(oJSON As Object) As Boolean
On Error GoTo ErrorHandler
    Dim bOk As Boolean
    Dim relation As Object
    Dim oProc As LDE.Procedure

    Dim ErrorMessage As String
    Dim warningmessage As String
    bOk = True

    sLog = sLog + Indent + "Adding relations..." + vbNewLine
    IncreaseIndent

    For Each relation In oJSON
    
        ErrorMessage = ""
        warningmessage = ""

        Set oProc = Database.Procedures("csp_lip_addRelations")
        oProc.Timeout = 299

        If Not oProc Is Nothing Then

            sLog = sLog + Indent + "Add relation between: " + relation.Item("table1") + "." + relation.Item("field1") + " and " + relation.Item("table2") + "." + relation.Item("field2") + vbNewLine

            oProc.Parameters("@@table1").InputValue = relation.Item("table1")
            oProc.Parameters("@@field1").InputValue = relation.Item("field1")
            oProc.Parameters("@@table2").InputValue = relation.Item("table2")
            oProc.Parameters("@@field2").InputValue = relation.Item("field2")

            Call oProc.Execute(False)

            ErrorMessage = oProc.Parameters("@@errorMessage").OutputValue
            warningmessage = oProc.Parameters("@@warningMessage").OutputValue
            
            If warningmessage <> "" Then
                sLog = sLog + Indent + (warningmessage) + vbNewLine
            End If
            
            'If errormessage is set, something went wrong
            If ErrorMessage <> "" Then
                sLog = sLog + Indent + (ErrorMessage) + vbNewLine
                bOk = False
            Else
                sLog = sLog + Indent + ("Relation between: " + relation.Item("table1") + "." + relation.Item("field1") + " and " + relation.Item("table2") + "." + relation.Item("field2") + " created.") + vbNewLine
            End If
            
            DecreaseIndent

        Else
            bOk = False
            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_addRelations'. Please make sure this procedure exists in the database and restart LDC.")
        End If

    Next relation
    DecreaseIndent

    Set oProc = Nothing
    
    InstallRelations = bOk

    Exit Function
ErrorHandler:
    Set oProc = Nothing
    InstallRelations = False
    Call UI.ShowError("lip.InstallRelations")
End Function

Private Function RollbackFieldsAndTables(sCreatedTables As String, sCreatedFields As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim oProc As New LDE.Procedure
    Set oProc = Database.Procedures("csp_lip_removeTablesAndFields")
    oProc.Timeout = 299
    
    If Not oProc Is Nothing Then
        If sCreatedFields <> "" Then
            Dim fieldArray() As String
            fieldArray() = VBA.Split(sCreatedFields, ";")
            
            For i = UBound(fieldArray) - 1 To LBound(fieldArray) Step -1
                oProc.Parameters("@@idfield") = CLng(fieldArray(i))
                Call oProc.Execute(False)
            Next i
        End If
        
        If sCreatedTables <> "" Then
            Dim tableArray() As String
            tableArray() = VBA.Split(sCreatedTables, ";")
            For i = UBound(tableArray) - 1 To LBound(tableArray) Step -1
                oProc.Parameters("@@idtable") = CLng(tableArray(i))
                Call oProc.Execute(False)
            Next i
        End If
    Else
        Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_removeTablesAndFields'. Please make sure this procedure exists in the database and restart LDC.")
        RollbackFieldsAndTables = False
        Exit Function
    End If
    
    RollbackFieldsAndTables = True
Exit Function
ErrorHandler:
    Call UI.ShowError("lip.RollbackFieldsAndTables")
End Function

'LJE 20160212 Check if a new version of LIP exists
Public Sub UpdateLIPOnNewVersion()
On Error GoTo ErrorHandler
    Dim Package As Object
    Dim PackageVersion As Double
    Dim downloadURL As String
    Dim InstallPath As String
    Dim PackageName As String
    
    Dim oPackageFile As Object
    Set oPackageFile = ReadPackageFile
    
    IndentLenght = "  "
    
    PackageName = "lip"
    Debug.Print Indent + "Checking version for LIP"
    Set Package = SearchForPackageInStores("lip")
    
    If Package Is Nothing Then
        Exit Sub
    End If
   
    PackageVersion = findNewestVersion(Package.Item("versions"))
    If PackageVersion > CDbl(VBA.Replace(oPackageFile.Item("lipversion"), ".", ",")) Then
        Debug.Print Indent + "Newer version of lip found"
        'LJE TODO Send in the versions
        Dim VBComps As Object
        Dim Path As String
        Dim tempModuleName As String
        
        Set VBComps = Application.VBE.ActiveVBProject.VBComponents
        'LJE TEST
        'VBComps.Item("lip").Name = "lip_old"
        'Call Application.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\LocalStore\lip\Install\VBA\lip.bas")
        
        'LJE TODO Update packages.json with new version
        oPackageFile.Item("lipversion") = VBA.Replace(PackageVersion, ",", ".")
        'LJE TEST
        'Call lip.RemoveModule("lip_old")
        Debug.Print Indent + "LIP updated"
    End If
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.UpdateLIPOnNewVersion")
End Sub
'LJE 20160212 Upgrade LIP if new version exists
Private Sub UpdateLIP()
On Error GoTo ErrorHandler
'Q: How to handle the remove of lip.bas.
'Separate lip functions in separate modules, an interface with functions which calls another bas which can be replaced.

'1. Replace lip.bas
'2. Replace csp (this is done manually now)
'3. Tell user what happened and what needs to be done.

 Dim VBComps As Object
 Dim Path As String
 Dim tempModuleName As String

 Set VBComps = Application.VBE.ActiveVBProject.VBComponents
 VBComps.Item("lip").Name = "lip_old"

 Call Application.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\LocalStore\lip\Install\VBA\lip.bas")
 
 'LJE TODO Update packages.json with new version

 Call lip.RemoveModule("lip_old")

'Call VBComps.Remove(VBComps.Item(tempModuleName)
 Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.UpdateLIP")
End Sub

'LJE Remove temporary lip.bas after update
Private Sub RemoveModule(sModuleName As String)
Dim VBComps As Object
On Error GoTo ErrorHandler

Set VBComps = Application.VBE.ActiveVBProject.VBComponents

Call VBComps.Remove(VBComps.Item(sModuleName))
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.RemoveModule")
End Sub

'LJE TODO Refactor with helper method to write json
Public Sub SetLipVersionInPackageFile(sVersion As String)
On Error GoTo ErrorHandler
'    Open ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "package.json" For Input As #1
'
'            ElseIf VBA.Dir(ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "app.json") <> "" Then
'                Open ThisApplication.WebFolder & DefaultInstallPath & PackageName & "\" & "app.json" For Input As #1
'
'            Else
'                Debug.Print (Indent + "Installation failed: couldn't find any package.json or app.json in the zip-file")
'                Exit Sub
'            End If
'
'            Do Until EOF(1)
'                Line Input #1, sLine
'                sJSON = sJSON & sLine
'            Loop
'
'            Close #1
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.SetLipVersionInPackageFile")
End Sub


