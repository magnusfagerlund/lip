Attribute VB_Name = "lip"

Option Explicit

'Lundalogik Package Store, DO NOT CHANGE, used to download system files for LIP
'Please add your own stores in packages.json
Private Const BaseURL As String = "http://limebootstrap.lundalogik.com"
Private Const ApiURL As String = "/api/apps/"

Private IndentLenght As String
Private Indent As String

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
Public Sub Install(PackageName As String, Optional Upgrade As Boolean)
On Error GoTo ErrorHandler
    Dim Package As Object
    Dim PackageVersion As Double
    Dim source As String

    IndentLenght = "  "
    
    'Check if first use ever
    If Dir(WebFolder + "packages.json") = "" Then
        Debug.Print "No packages.json found, assuming fresh install"
        Call InstallLIP
    End If
    
    PackageName = LCase(PackageName)
    
    Debug.Print "====== LIP Install: " + PackageName + " ======"
    
    Debug.Print "Looking for package: '" + PackageName + "'"
    Set Package = SearchForPackageOnStores(PackageName)
    If Package Is Nothing Then
        Exit Sub
    End If
    
    If Package.Exists("source") Then
        source = VBA.Replace(Package.Item("source"), "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with URLs
    Else
        source = BaseURL & ApiURL 'Use Lundalogik Packagestore if source-node wasn't found
    End If
    
    Set Package = Package.Item("info")
    
    'Parse result from store
    PackageVersion = findNewestVersion(Package.Item("versions"))
    
    'Check if package already exsists
    If Not Upgrade Then
        If CheckForLocalInstalledPackage(PackageName, PackageVersion) = True Then
            Exit Sub
        End If
    End If
    
    'Install dependecies
    If Package.Exists("dependencies") Then
        IncreaseIndent
        Call InstallDependencies(Package)
        DecreaseIndent
    End If
    
    'Download and unzip
    Debug.Print "Downloading '" + PackageName + "' files..."
    
    Call DownloadFile(PackageName, source)
    Call Unzip(PackageName)
    Debug.Print "Download complete!"
   
    Call InstallPackageComponents(PackageName, PackageVersion, Package)
    
    Debug.Print "==================================="
    
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.Install")
End Sub

'Installs package from a zip-file. Input parameter: complete searchpath to the zip-file, including the filename
Public Sub InstallFromZip(ZipPath As String)
On Error GoTo ErrorHandler

    'Check if valid path
    If VBA.Right(ZipPath, 4) = ".zip" Then
        If VBA.Dir(ZipPath) <> "" Then
            'Check if first use ever
            If Dir(WebFolder + "packages.json") = "" Then
                Debug.Print "No packages.json found, assuming fresh install"
                Call InstallLIP
            End If
            
'           Copy file to actionpads\apps
            Dim PackageName As String
            Dim strArray() As String
            strArray = VBA.Split(ZipPath, "\")
            PackageName = VBA.Split(strArray(UBound(strArray)), ".")(0)
            Debug.Print "====== LIP Install: " + PackageName + " ======"
            Debug.Print "Copying and unzipping file"
            
            'Copy zip-file to the apps-folder if it's not already there
            If ZipPath <> ThisApplication.WebFolder & "apps\" & PackageName & ".zip" Then
                Call VBA.FileCopy(ZipPath, ThisApplication.WebFolder & "apps\" & PackageName & ".zip")
            End If
            
'           Unzip file
            Call Unzip(PackageName) 'Filename without fileextension as parameter
            
            'Get package information from json-file
            Dim Package As Object
            Dim sJSON As String
            Dim sLine As String
            
            Open ThisApplication.WebFolder & "apps\" & PackageName & "\" & "app.json" For Input As #1
            'TODO: Catch if app.json is missing
            
            Do Until EOF(1)
                Line Input #1, sLine
                sJSON = sJSON & sLine
            Loop
            
            Close #1
            
            Set Package = JSON.parse(sJSON)
            
            'Install dependencies
            If Package.Exists("dependencies") Then
                IncreaseIndent
                Call InstallDependencies(Package)
                DecreaseIndent
            End If
            
            Call InstallPackageComponents(PackageName, 1, Package)
    
            Debug.Print "==================================="
        Else
            Debug.Print ("Couldn't find file.")
        End If
    Else
        Debug.Print ("Path must end with .zip")
    End If


Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallFromZip")
End Sub

'Installs all packages defined in the packages.json file
Public Sub InstallFromPackageFile()
On Error GoTo ErrorHandler
    Dim LocalPackages As Object
    Dim LocalPackageName As Variant
    
    Debug.Print "Installing dependecies from packages.json file..."
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


Private Sub InstallPackageComponents(PackageName As String, PackageVersion As Double, Package)
On Error GoTo ErrorHandler

    
    'Install localizations
    If Package.Item("install").Exists("localize") = True Then
        Debug.Print Indent + "Adding localizations..."
        IncreaseIndent
        Call InstallLocalize(Package.Item("install").Item("localize"))
        DecreaseIndent
          
    End If
    
    'Install VBA
    If Package.Item("install").Exists("vba") = True Then
        Debug.Print Indent + "Adding VBA modules, forms and classes..."
        IncreaseIndent
        Call InstallVBAComponents(PackageName, Package.Item("install").Item("vba"))
        DecreaseIndent
    End If
    
    If Package.Item("install").Exists("tables") = True Then
        IncreaseIndent
        Call InstallFieldsAndTables(Package.Item("install").Item("tables"))
        DecreaseIndent
    End If
    
    If Package.Item("install").Exists("sql") = True Then
        IncreaseIndent
        Call InstallSQL(Package.Item("install").Item("sql"), PackageName)
        DecreaseIndent
    End If
    'Update packages.json
    Call WriteToPackageFile(PackageName, CStr(PackageVersion))
    
    Debug.Print Indent + "Installation of " + PackageName + " done!"
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallPackageComponents")
End Sub

Private Sub InstallDependencies(Package As Object)
On Error GoTo ErrorHandler
    Dim DependencyName As Variant
    Dim LocalPackage As Object
    Debug.Print Indent + "Dependencies found! Installing..."
    IncreaseIndent
    For Each DependencyName In Package.Item("dependencies").keys()
        Set LocalPackage = FindPackageLocally(CStr(DependencyName))
        If LocalPackage Is Nothing Then
            Debug.Print Indent + "Installing dependency: " + CStr(DependencyName)
            Call Install(CStr(DependencyName))
        ElseIf CDbl(VBA.Replace(LocalPackage.Item(DependencyName), ".", ",")) < CDbl(VBA.Replace(Package.Item("dependencies").Item(DependencyName), ".", ",")) Then
            Call Install(CStr(DependencyName), True)
        Else
        End If
    Next DependencyName
    DecreaseIndent
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallDependencies")
End Sub


Private Function SearchForPackageOnStores(PackageName As String) As Object
On Error GoTo ErrorHandler
    Dim sJSON As String
    Dim oJSON As Object
    Dim oPackages As Object
    Dim Path As String
    Dim oPackage As Variant

    Set oPackages = ReadPackageFile.Item("stores")

    'Loop through packagestores from packages.json
    For Each oPackage In oPackages

        Path = oPackages.Item(oPackage)
        Debug.Print ("Looking for package at store '" & oPackage & "'")
        sJSON = getJSON(Path + PackageName + "/")
        
        If sJSON <> "" Then
            sJSON = VBA.Left(sJSON, VBA.Len(sJSON) - 1) & ",""source"":""" & oPackages.Item(oPackage) & """}" 'Add a source node so we know where the package exists
        End If
        
        Set oJSON = parseJSON(sJSON) 'Create a JSON object from the string

        If Not oJSON Is Nothing Then
            If oJSON.Item("error") = "" Then
                'Package found, make sure the install node exists
                If Not oJSON.Item("info").Item("install") Is Nothing Then
                    Debug.Print ("Package '" & PackageName & "' found on store '" & oPackage & "'")
                    Set SearchForPackageOnStores = oJSON
                    Exit Function
                Else
                    Debug.Print ("Package '" & PackageName & "' found on store '" & oPackage & "' but has no valid install instructions!")
                    Set SearchForPackageOnStores = Nothing
                    Exit Function
                End If
            End If
        End If
    Next

    'If we've reached this code, package wasn't found
    Debug.Print ("Package '" & PackageName & "' not found!")
    Set SearchForPackageOnStores = Nothing

Exit Function
ErrorHandler:
    Set SearchForPackageOnStores = Nothing
    Call UI.ShowError("lip.SearchForPackageOnStores")
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
            Debug.Print "Current version of" + PackageName + " is already installed, please use the upgrade command to reinstall package"
            Debug.Print "==================================="
            CheckForLocalInstalledPackage = True
            Exit Function
        ElseIf PackageVersion > LocalPackageVersion Then
            Debug.Print "Package " + PackageName + " is already installed, please use the upgrade command to upgrade package from " + Format(LocalPackageVersion, "0.0") + " -> " + Format(PackageVersion, "0.0")
            Debug.Print "==================================="
            CheckForLocalInstalledPackage = True
            Exit Function
        Else
            Debug.Print "A newer version of " + PackageName + " is already installed. Remote: " + Format(PackageVersion, "0.0") + " ,Local: " + Format(LocalPackageVersion, "0.0") + ". Please use the upgrade command to reinstall package"
            Debug.Print "==================================="
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

Private Function parseJSON(sJSON As String) As Object
On Error GoTo ErrorHandler
    Dim oJSON As Object
    Set oJSON = JSON.parse(sJSON)
    Set parseJSON = oJSON
Exit Function
ErrorHandler:
    Set parseJSON = Nothing
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

Private Sub InstallLocalize(oJSON As Object)
On Error GoTo ErrorHandler
    Dim Localize As Variant
        
    For Each Localize In oJSON
        Call AddOrCheckLocalize( _
            Localize.Item("owner"), _
            Localize.Item("context"), _
            "", _
            Localize.Item("en-us"), _
            Localize.Item("sv"), _
            Localize.Item("no"), _
            Localize.Item("fi") _
        )
    Next Localize
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallLocalize")
End Sub

Private Sub InstallSQL(oJSON As Object, PackageName As String)
On Error GoTo ErrorHandler
    Dim Sql As Variant
    Dim oProc As New LDE.Procedure
    Dim strSQL As String
    Dim sLine As String
    Dim sErrormessage As String
        
    Debug.Print Indent + "Installing SQL..."
    IncreaseIndent
    For Each Sql In oJSON
    
        strSQL = ""
        sErrormessage = ""

        Open ThisApplication.WebFolder & "apps\" & PackageName & "\" & Sql.Item("relPath") For Input As #1
            Do Until EOF(1)
                Line Input #1, sLine
                strSQL = strSQL & sLine & vbNewLine
            Loop
            Close #1
            
            Set oProc = Database.Procedures("csp_lip_installSQL")
            If Not oProc Is Nothing Then
                oProc.Parameters("@@sql") = strSQL
                oProc.Parameters("@@name") = Sql.Item("name")
                oProc.Parameters("@@type") = Sql.Item("type")
                oProc.Execute (False)
                
                sErrormessage = oProc.Parameters("@@errormessage").OutputValue
                
                If sErrormessage <> "" Then
                    Debug.Print (sErrormessage)
                Else
                    Debug.Print ("'" & Sql.Item("name") & "'" & " added.")
                End If
                
            Else
                Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_installSQL'. Please make sure this procedure exists in the database and restart LDC.")
            End If
            
    Next Sql
    DecreaseIndent
Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallSQL")
End Sub

Private Sub InstallFieldsAndTables(oJSON As Object)
On Error GoTo ErrorHandler
    Dim table As Object
    Dim oProc As LDE.Procedure
    Dim field As Object
    Dim idtable As Long
    Dim iddescriptiveexpression As Long
    Dim oItem As Variant
    
    Dim localname_singular As String
    Dim localname_plural As String
    Dim errormessage As String
    
    Debug.Print "Adding fields and tables..."
    IncreaseIndent
    
    For Each table In oJSON
        localname_singular = ""
        localname_plural = ""
        errormessage = ""
        
        Set oProc = Database.Procedures("csp_lip_createtable")
        
        If Not oProc Is Nothing Then
        
            Debug.Print Indent + "Add table: " + table.Item("name")
            
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
            
            errormessage = oProc.Parameters("@@errorMessage").OutputValue
            
            idtable = oProc.Parameters("@@idtable").OutputValue
            iddescriptiveexpression = oProc.Parameters("@@iddescriptiveexpression").OutputValue
            
            'If errormessage is set, something went wrong
            If errormessage <> "" Then
                Debug.Print (errormessage)
            Else
                Debug.Print ("Table """ & table.Item("name") & """ created.")
            End If
            
            ' Create fields
            IncreaseIndent
            If table.Exists("fields") Then
                For Each field In table.Item("fields")
                    Debug.Print Indent + "Add field: " + field.Item("name")
                    Call AddField(table.Item("name"), field)
                Next field
            End If
            
            'Set table attributes(must be done AFTER fields has been created in order to be able to set descriptive expression)
            'Only set attributes if table was created
            If idtable <> -1 Then
                Call SetTableAttributes(table, idtable, iddescriptiveexpression)
            End If
            
            DecreaseIndent
            
        Else
            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_createtable'. Please make sure this procedure exists in the database and restart LDC.")
        End If
        
    Next table
    DecreaseIndent
    
    Set oProc = Nothing
    
    Exit Sub
ErrorHandler:
    Set oProc = Nothing
    Call UI.ShowError("lip.InstallFieldsAndTables")
End Sub


Private Sub AddField(tableName As String, field As Object)
On Error GoTo ErrorHandler
    Dim oProc As New LDE.Procedure
    Dim errormessage As String
    Dim fieldLocalnames As String
    Dim separatorLocalnames As String
    Dim oItem As Variant
    errormessage = ""
    fieldLocalnames = ""
    separatorLocalnames = ""
    Set oProc = Database.Procedures("csp_lip_createfield")
    
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
                        Debug.Print ("No support for setting field attribute " & oItem)
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
        
        Call oProc.Execute(False)
        
        errormessage = oProc.Parameters("@@errorMessage").OutputValue
        
        'If errormessage is set, something went wrong
        If errormessage <> "" Then
            Debug.Print (errormessage)
        Else
            Debug.Print ("Field """ & field.Item("name") & """ created.")
        End If
    Else
        Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_createfield'. Please make sure this procedure exists in the database and restart LDC.")
    End If
    Set oProc = Nothing
    
    Exit Sub
ErrorHandler:
    Set oProc = Nothing
    Call UI.ShowError("lip.AddField")
End Sub

Private Sub SetTableAttributes(ByRef table As Object, idtable As Long, iddescriptiveexpression As Long)
On Error GoTo ErrorHandler

    Dim oProcAttributes As LDE.Procedure
    Dim oItem As Variant
    Dim errormessage As String
    
    If table.Exists("attributes") Then
    
        Set oProcAttributes = Application.Database.Procedures("csp_lip_settableattributes")
        
        If Not oProcAttributes Is Nothing Then
        
            Debug.Print Indent + "Adding attributes for table: " + table.Item("name")
        
            oProcAttributes.Parameters("@@tablename").InputValue = table.Item("name")
            oProcAttributes.Parameters("@@idtable").InputValue = idtable
            oProcAttributes.Parameters("@@iddescriptiveexpression").InputValue = iddescriptiveexpression
        
            For Each oItem In table.Item("attributes")
                If oItem <> "" Then
                    If Not oProcAttributes.Parameters.Lookup("@@" & oItem, lkLookupProcedureParameterByName) Is Nothing Then
                        oProcAttributes.Parameters("@@" & oItem).InputValue = table.Item("attributes").Item(oItem)
                    Else
                        Debug.Print ("No support for setting table attribute " & oItem)
                    End If
                End If
            Next
            
            Call oProcAttributes.Execute(False)
        
            errormessage = oProcAttributes.Parameters("@@errorMessage").OutputValue
            
            'If errormessage is set, something went wrong
            If errormessage <> "" Then
                Debug.Print (errormessage)
            Else
                Debug.Print ("Attributes for table """ & table.Item("name") & """ set.")
            End If
        
        Else
            Call Lime.MessageBox("Couldn't find SQL-procedure 'csp_lip_settableattributes'. Please make sure this procedure exists in the database and restart LDC.")
        End If
    End If
    
    Set oProcAttributes = Nothing
    
    Exit Sub
ErrorHandler:
    Set oProcAttributes = Nothing
    Call UI.ShowError("lip.SetTableAttributes")
End Sub

Private Sub DownloadFile(PackageName As String, Path As String)
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
    If WinHttpReq.status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile WebFolder + "apps\" + PackageName + ".zip", 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.DownloadFile")
End Sub

Private Sub Unzip(PackageName)
On Error GoTo ErrorHandler
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = WebFolder + "apps\" + PackageName + ".zip"
    FileNameFolder = WebFolder & "apps\" & PackageName & "\"

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

Private Sub InstallVBAComponents(PackageName As String, VBAModules As Object)
On Error GoTo ErrorHandler
    Dim VBAModule As Variant
    IncreaseIndent
    For Each VBAModule In VBAModules
        Call addModule(PackageName, VBAModule.Item("name"), VBAModule.Item("relPath"))
        Debug.Print Indent + "Added " + VBAModule.Item("name")
    Next VBAModule
    DecreaseIndent
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallVBAComponents")
End Sub

Private Sub addModule(PackageName As String, ModuleName As String, RelPath As String)
On Error GoTo Errorhandler
    If PackageName <> "" And ModuleName <> "" Then
        Dim VBComps As Object
        Dim Path As String
        Dim tempModuleName As String
        
        Set VBComps = Application.VBE.ActiveVBProject.VBComponents
        If ComponentExists(ModuleName, VBComps) = True Then
            tempModuleName = LCO.GenerateGUID
            tempModuleName = VBA.Replace(VBA.Mid(tempModuleName, 2, VBA.Len(tempModuleName) - 2), "-", "")
            tempModuleName = VBA.Left("temp" & tempModuleName, 30)
            VBComps.Item(ModuleName).Name = tempModuleName
            Call VBComps.Remove(VBComps.Item(tempModuleName))
        End If
        Path = WebFolder + "apps\" + PackageName + "\" + RelPath
     
        Call Application.VBE.ActiveVBProject.VBComponents.Import(Path)
    End If
    Exit Sub
Errorhandler:
    Call UI.ShowError("lip.addModule")
End Sub

Private Function ComponentExists(ComponentName As String, VBComps As Object) As Boolean
On Error GoTo ErrorHandler
    Dim VBComp As Variant
    
    For Each VBComp In VBComps
        If VBComp.name = ComponentName Then
             ComponentExists = True
             Exit Function
        End If
    Next VBComp
    
    ComponentExists = False
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("lip.ComponentExists")
End Function

Private Sub WriteToPackageFile(PackageName As String, Version As String)
On Error GoTo ErrorHandler
    Dim oJSON As Object
    Dim fs As Object
    Dim a As Object
    Dim Line As Variant
    Set oJSON = ReadPackageFile
    
    oJSON.Item("dependencies").Item(PackageName) = Version

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WebFolder + "packages.json", True)
    For Each Line In Split(PrettyPrintJSON(JSON.toString(oJSON)), vbCrLf)
        Line = VBA.Replace(Line, "\/", "/") 'Replace \/ with only / since JSON escapes frontslash with a backslash which causes problems with packagestores URLs
        a.WriteLine Line
    Next Line
    a.Close
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.WriteToPackageFile")
End Sub

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
        Debug.Print "Error: No packages.json found!"
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
            Debug.Print ("Couldn't find dependencies in packages.json")
        End If
        
    End If
    
    Set FindPackageLocally = Nothing
    Exit Function
ErrorHandler:
    Set FindPackageLocally = Nothing
    Call UI.ShowError("lip.FindPackageLocally")
End Function

Private Sub CreateANewPackageFile()
On Error GoTo ErrorHandler
    Dim fs As Object
    Dim a As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(WebFolder + "packages.json", True)
    a.WriteLine ("{")
    a.WriteLine ("    ""stores"":{")
    a.WriteLine ("        ""PackageStore"":""http://limebootstrap.lundalogik.com/api/apps/"",")
    a.WriteLine ("        ""Bootstrap Appstore"":""http://limebootstrap.lundalogik.com/api/apps/""")
    a.WriteLine ("    },")
    a.WriteLine ("    ""dependencies"":{")
    a.WriteLine ("    }")
    a.WriteLine ("}")
    a.Close
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.CreateNewPackageFile")
End Sub

Public Sub InstallLIP()
On Error GoTo ErrorHandler

    Debug.Print "Creating a new packages.json file..."
    Call CreateANewPackageFile
    
    Debug.Print "Installing JSON-lib..."
    Call DownloadFile("vba_json", BaseURL + ApiURL)
    Call Unzip("vba_json")
    Call addModule("vba_json", "JSON", "JSON.bas")
    Call addModule("vba_json", "cStringBuilder", "cStringBuilder.cls")
    
    Call WriteToPackageFile("vba_json", "1")

    Debug.Print "Install of LIP complete!"
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.InstallLIP")
End Sub

Private Function AddOrCheckLocalize(sOwner As String, sCode As String, sDescription As String, sEN_US As String, sSV As String, sNO As String, sFI As String) As Boolean
On Error GoTo ErrorHandler
    Dim oFilter As New LDE.Filter
    Dim oRecs As New LDE.Records
    
    Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    Call oFilter.AddCondition("code", lkOpEqual, sCode)
    oFilter.AddOperator lkOpAnd
    
    If oFilter.HitCount(Database.Classes("localize")) = 0 Then
        Debug.Print (Indent + "Localization " & sOwner & "." & sCode & " not found, creating new!")
        Dim oRec As New LDE.Record
        Call oRec.Open(Database.Classes("localize"))
        oRec.Value("owner") = sOwner
        oRec.Value("code") = sCode
        oRec.Value("context") = sDescription
        oRec.Value("sv") = sSV
        oRec.Value("en-us") = sEN_US
        oRec.Value("no") = sNO
        oRec.Value("fi") = sFI
        Call oRec.Update
    ElseIf oFilter.HitCount(Database.Classes("localize")) = 1 Then
    Debug.Print (Indent + "Updating localization " & sOwner & "." & sCode)
        Call oRecs.Open(Database.Classes("localize"), oFilter)
        oRecs(1).Value("owner") = sOwner
        oRecs(1).Value("code") = sCode
        oRecs(1).Value("context") = sDescription
        oRecs(1).Value("sv") = sSV
        oRecs(1).Value("en-us") = sEN_US
        oRecs(1).Value("no") = sNO
        oRecs(1).Value("fi") = sFI
        Call oRecs.Update
        
    Else
        Call MsgBox("There are multiple copies of " & sOwner & "." & sCode & "  which is bad! Fix it", vbCritical, "To many translations makes Jack a dull boy")
    End If
    
    Set Localize.dicLookup = Nothing
    AddOrCheckLocalize = True
    Exit Function
ErrorHandler:
    Debug.Print ("Error while validating or adding Localize")
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
    Indent = Left(Indent, Len(Indent) - Len(IndentLenght))
    Exit Sub
ErrorHandler:
    Call UI.ShowError("lip.DecreaseIndent")
End Sub
