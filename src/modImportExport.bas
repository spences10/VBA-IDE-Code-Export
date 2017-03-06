Attribute VB_Name = "modImportExport"
Option Explicit

'// Updates the configuration file for the current active project.
'// * Entries for modules not yet declared in the configuration file as created.
'// * Modules listed in the configuration file which are not found are prompted
'//   to be deleted from the configuration file.
'// * The current loaded references are used to update the configuration file.
'// * References in the configuration file whic hare not loaded are prompted to
'//   be deleted from the configuration file.
Public Sub MakeConfigFile()

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration

    Dim comModule           As VBComponent
    Dim boolDeleteModule    As Boolean
    Dim boolCreateNewEntry  As Boolean
    Dim varModuleName       As Variant
    Dim strModuleName       As String

    Dim refReference        As Reference
    Dim lngIndex            As Long
    Dim varIndex            As Variant
    Dim boolForbiddenRef    As Boolean

    Dim collDeleteList      As Collection
    Dim strDeleteListStr    As String
    Dim intUserResponse     As Integer

    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Generate entries for modules not yet listed
    For Each comModule In prjActProj.VBComponents
        boolCreateNewEntry = _
            ExportableModule(comModule) And _
            Not Config.ModuleDeclared(comModule.Name)

        If boolCreateNewEntry Then
            Config.ModulePath(comModule.Name) = comModule.Name & "." & FileExtension(comModule)
        End If
    Next comModule

    '// Ask user if they want to delete entries for missing modules
    Set collDeleteList = New Collection
    strDeleteListStr = ""
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        boolDeleteModule = True
        If CollectionKeyExists(prjActProj.VBComponents, strModuleName) Then
            If ExportableModule(prjActProj.VBComponents(strModuleName)) Then
                boolDeleteModule = False
            End If
        End If
        If boolDeleteModule Then
            collDeleteList.Add strModuleName
            strDeleteListStr = strDeleteListStr & strModuleName & vbNewLine
        End If
    Next varModuleName

    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
            Prompt:= _
                "There are some modules listed in the configuration file which " & _
                "haven't been found in the current project. Would you like to " & _
                "remove these modules from the configuration file?" & vbNewLine & _
                vbNewLine & _
                "Missing modules:" & vbNewLine & _
                strDeleteListStr, _
            Buttons:=vbYesNo + vbDefaultButton2, _
            Title:="Missing Modules")

        If intUserResponse = vbYes Then
            For Each varModuleName In collDeleteList
                strModuleName = varModuleName
                Config.ModulePathRemove strModuleName
            Next varModuleName
        End If
    End If

    '// Generate entries for references in the current VBProject
    For Each refReference In prjActProj.References
        If Not refReference.BuiltIn Then
            boolForbiddenRef = _
                refReference.Name = "stdole" Or _
                refReference.Name = "Office" Or _
                refReference.Name = "MSForms"
            If Not boolForbiddenRef Then
                Config.ReferencesUpdateFromVBRef refReference
            End If
        End If
    Next refReference

    '// Prompt if entries for missing references should be deleted
    Set collDeleteList = New Collection
    strDeleteListStr = ""
    For lngIndex = Config.ReferencesCount To 1 Step -1
        If Not CollectionKeyExists(prjActProj.References, Config.ReferenceName(lngIndex)) Then
            collDeleteList.Add lngIndex
            strDeleteListStr = strDeleteListStr & Config.ReferenceName(lngIndex) & vbNewLine
        End If
    Next

    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
            Prompt:= _
                "There are some references listed in the configuration file which " & _
                "haven't been found in the current project. Would you like to " & _
                "remove these references from the configuration file?" & vbNewLine & _
                vbNewLine & _
                "Missing references:" & vbNewLine & _
                strDeleteListStr, _
            Buttons:=vbYesNo + vbDefaultButton2, _
            Title:="Missing References")

        If intUserResponse = vbYes Then
            For Each varIndex In collDeleteList
                lngIndex = varIndex
                Config.ReferenceRemove lngIndex
            Next varIndex
        End If
    End If

    '// Write changes to config file
    Config.WriteToProjectConfigFile

    MsgBox _
        "Configuration file was successfully updated. Please review the " & _
        "file with a text editor."

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub

'// Exports code modules and cleans the current active VBProject as specified
'// by the project's configuration file.
'// * Any code module in the VBProject which is listed in the configuration
'//   file is exported to the configured path.
'// * code modules which were exported are deleted or cleared.
'// * References loaded in the Project which are listed in the configuration
'//   file is deleted.
Public Sub Export()

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration
    Dim comModule           As VBComponent
    Dim lngIndex            As Long
    Dim strModuleName       As String
    Dim varModuleName       As Variant

    On Error GoTo ErrHandler

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Export all modules listed in the configuration
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        ' TODO Provide a warning if module listed in configuration is not found
        If CollectionKeyExists(prjActProj.VBComponents, strModuleName) Then
            Set comModule = prjActProj.VBComponents(strModuleName)
            EnsurePath Config.ModuleFullPath(strModuleName)
            comModule.Export Config.ModuleFullPath(strModuleName)

            If comModule.Type = vbext_ct_Document Then
                comModule.CodeModule.DeleteLines 1, comModule.CodeModule.CountOfLines
            Else
                prjActProj.VBComponents.Remove comModule
            End If
        End If
    Next varModuleName

    '// Remove all references listed
    For lngIndex = 1 To Config.ReferencesCount
        If CollectionKeyExists(prjActProj.References, Config.ReferenceName(lngIndex)) Then
            prjActProj.References.Remove prjActProj.References(Config.ReferenceName(lngIndex))
        End If
    Next lngIndex

exitSub:
    Exit Sub

ErrHandler:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub

'// Imports textual data from the file system such as VBA code to build the
'// current active VBProject as specified in it's configuration file.
'// * Each code module file listed in the configuration file is imported into
'//   the VBProject. Modules with the same name are overwritten.
'// * All references declared in the configuration file are loaded into the
'//   project.
'// * The project name is set to the project name specified by the configuration
'//   file.
Public Sub Import()

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration
    Dim strModuleName       As String
    Dim varModuleName       As Variant

    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Import code from listed module files
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        ImportModule prjActProj, strModuleName, Config.ModuleFullPath(strModuleName)
    Next varModuleName

    '// Add references listed in the config file
    Config.ReferencesAddToVBRefs prjActProj.References

    '// Set the VBA Project name
    If Config.VBAProjectNameDeclared Then
        prjActProj.Name = Config.VBAProjectName
    End If

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub


'// Import a VBA code module... how hard could it be right?
Private Sub ImportModule(ByVal Project As VBProject, ByVal moduleName As String, ByVal ModulePath As String)

    Dim comNewImport        As VBComponent
    Dim comExistingComp     As VBComponent
    Dim modCodeCopy         As CodeModule
    Dim modCodePaste        As CodeModule

    Set comNewImport = Project.VBComponents.Import(ModulePath)
    If comNewImport.Name <> moduleName Then
        If CollectionKeyExists(Project.VBComponents, moduleName) Then

            Set comExistingComp = Project.VBComponents(moduleName)
            If comExistingComp.Type = vbext_ct_Document Then

                Set modCodeCopy = comNewImport.CodeModule
                Set modCodePaste = comExistingComp.CodeModule
                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.Lines(1, modCodeCopy.CountOfLines)
                End If
                Project.VBComponents.Remove comNewImport

            Else

                Project.VBComponents.Remove comExistingComp
                comNewImport.Name = moduleName

            End If
        Else

            comNewImport.Name = moduleName

        End If
    End If

End Sub


'// Is the given module exportable by this tool?
Private Function ExportableModule(ByVal comModule As VBComponent) As Boolean

    ExportableModule = _
        (Not ModuleEmpty(comModule)) And _
        (Not FileExtension(comModule) = vbNullString)

End Function


'// Check if a code module is effectively empty.
'// effectively empty should be functionally and semantically equivelent to
'// actually empty.
Private Function ModuleEmpty(ByVal comModule As VBComponent) As Boolean

    Dim lngNumLines As Long
    Dim lngCurLine As Long
    Dim strCurLine As String

    ModuleEmpty = True

    lngNumLines = comModule.CodeModule.CountOfLines
    For lngCurLine = 1 To lngNumLines
        strCurLine = comModule.CodeModule.Lines(lngCurLine, 1)
        If Not (strCurLine = "Option Explicit" Or strCurLine = "") Then
            ModuleEmpty = False
            Exit Function
        End If
    Next lngCurLine

End Function


'// The appropriate file extension for exporting the given module
Private Function FileExtension(ByVal comModule As VBComponent) As String

    Select Case comModule.Type
        Case vbext_ct_StdModule
            FileExtension = "bas"
        Case vbext_ct_ClassModule, vbext_ct_Document
            FileExtension = "cls"
        Case vbext_ct_MSForm
            FileExtension = "frm"
        Case Else
            FileExtension = vbNullString
    End Select

End Function


'// Ensure path to a file exists. Creates missing folders.
Private Sub EnsurePath(ByVal Path As String)

    Dim strParentPath As String

    Set FSO = New Scripting.FileSystemObject
    strParentPath = FSO.GetParentFolderName(Path)

    If Not strParentPath = "" Then
        EnsurePath strParentPath
        If Not FSO.FolderExists(strParentPath) Then
            If FSO.FileExists(strParentPath) Then
                Err.Raise vbObjectError + 1, "modImportExport:EnsurePath", "A file exists where a folder needs to be: " & strParentPath
            Else
                FSO.CreateFolder (strParentPath)
            End If
        End If
    End If

End Sub
