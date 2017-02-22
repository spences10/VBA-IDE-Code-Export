VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ForReading                As Integer = 1
Private Const STRCONFIGFILENAME         As String = "CodeExport.config.json"

Private Const STR_CONFIGKEY_PROJECTNAME             As String = "VBAProject Name"
Private Const STR_CONFIGKEY_BASEPATH                As String = "Base Path"
Private Const STR_CONFIGKEY_MODULEPATHS             As String = "Module Paths"
Private Const STR_CONFIGKEY_REFERENCES              As String = "References"
Private Const STR_CONFIGKEY_REFERENCE_NAME          As String = "Name"
Private Const STR_CONFIGKEY_REFERENCE_DESCRIPTION   As String = "Description"
Private Const STR_CONFIGKEY_REFERENCE_GUID          As String = "GUID"
Private Const STR_CONFIGKEY_REFERENCE_MAJOR         As String = "Major"
Private Const STR_CONFIGKEY_REFERENCE_MINOR         As String = "Minor"
Private Const STR_CONFIGKEY_REFERENCE_PATH          As String = "Path"

Private pProject        As VBProject
Private pConfig         As Dictionary


Private Sub Class_Initialize()

    Set pConfig = New Dictionary
    Set pProject = Nothing

End Sub


'// The VBProject context for this configuration. Used for file paths
Public Property Get Project() As VBProject

    Set Project = pProject

End Property

Public Property Let Project(ByVal newProject As VBProject)

    Set pProject = newProject

End Property


'//
'// VBAProject name configuration
'//

'// Get the VBAProject name as determined by the configuration
Public Property Get VBAProjectName() As String

    If pConfig.Exists(STR_CONFIGKEY_PROJECTNAME) Then
        VBAProjectName = pConfig(STR_CONFIGKEY_PROJECTNAME)
    Else
        VBAProjectName = vbNullString
    End If

End Property

'// Set the configuration VBAProject name
Public Property Let VBAProjectName(ByVal newVBAProjectName As String)

    pConfig(STR_CONFIGKEY_PROJECTNAME) = newVBAProjectName

End Property

'// Check if the VBAProject name is actually declared in the configuration
'// This distinguishes between VBAProjectName = "" and not being set at all.
Public Property Get VBAProjectNameDeclared() As Boolean

    VBAProjectNameDeclared = pConfig.Exists(STR_CONFIGKEY_PROJECTNAME)

End Property

'// Remove the VBAProject name declaration from the configuration
'// This makes VBAProjectNameDeclare = false
Public Sub VBAProjectNameRemove()
    If pConfig.Exists(STR_CONFIGKEY_PROJECTNAME) Then
        pConfig.Remove STR_CONFIGKEY_PROJECTNAME
    End If
End Sub


'//
'// Base Path Configuration for prefixing all module paths
'//

'// Get the base path as determined by the configuration
Public Property Get BasePath() As String
    If pConfig.Exists(STR_CONFIGKEY_BASEPATH) Then
        BasePath = pConfig(STR_CONFIGKEY_BASEPATH)
    Else
        BasePath = vbNullString
    End If
End Property

'// Set the configuration base path
Public Property Let BasePath(ByVal newBasePath As String)
    pConfig(STR_CONFIGKEY_BASEPATH) = newBasePath
End Property

'// Check if the base path is actually declared in the configuration
'// This distinguishes between BasePath = "" and not being set at all.
Public Property Get BasePathDeclared() As Boolean
    BasePathDeclared = pConfig.Exists(STR_CONFIGKEY_BASEPATH)
End Property

'// Remove the base path declaration from the configuration
'// This makes BasePathDeclared = false
Public Sub BasePathRemove()
    If pConfig.Exists(STR_CONFIGKEY_BASEPATH) Then
        pConfig.Remove STR_CONFIGKEY_BASEPATH
    End If
End Sub


'//
'// Module paths configuration
'//

'// Check if a path for the given module is declared in the configuration
Public Property Get ModuleDeclared(ByVal moduleName As String) As Boolean
' TODO Rename to ModulePathDeclared
    ModuleDeclared = ModulePathsDictReadOnly.Exists(moduleName)

End Property

'// List all the names of the modules which have been assigned a path in the
'// configuration
Public Property Get ModuleNames() As Variant()

    ModuleNames = ModulePathsDictReadOnly.Keys

End Property

'// Get the module path for a module as determined by the configuration
Public Property Get ModulePath(ByVal moduleName As String) As String

    ModulePath = ModulePathsDictReadOnly(moduleName)

End Property

'// Set the configuration module path for a module
Public Property Let ModulePath(ByVal moduleName As String, ByVal newPath As String)

    ModulePathsDictForWriting(moduleName) = newPath

End Property

'// Evaluate a module's absolute module path as determined by the configuration
'// and the path of the project.
Public Property Get ModuleFullPath(ByVal moduleName As String) As String

    ModuleFullPath = FSO.BuildPath(BasePath, ModulePath(moduleName))
    If FSO.GetDriveName(ModuleFullPath) = vbNullString Then
        ModuleFullPath = FSO.BuildPath(ProjectDir, ModuleFullPath)
    End If
    ModuleFullPath = FSO.GetAbsolutePathName(ModuleFullPath)

End Property

'// Remove the module path declaration for a module from the configuration
'// This makes ModuleDeclared(moduleName) = false
Public Sub ModulePathRemove(ByVal moduleName As String)

    ModulePathsDictReadOnly.Remove moduleName

End Sub

'// Remove the entire Module Paths declaration from the configuration
Public Sub ModulePathsRemove()
    If pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        pConfig.Remove STR_CONFIGKEY_MODULEPATHS
    End If
End Sub

'// Get a dictionary the represents the module paths configuration
'// Note that it may not actually be the configuration dictionary so don't write
'// to it.
Private Property Get ModulePathsDictReadOnly() As Dictionary

    If pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        Set ModulePathsDictReadOnly = pConfig(STR_CONFIGKEY_MODULEPATHS)
    Else
        Set ModulePathsDictReadOnly = New Dictionary
    End If

End Property

'// Get the dictionary of the module paths configuration.
'// If the module paths is not yet declared, this will create an empty
'// declaration
Private Property Get ModulePathsDictForWriting() As Dictionary

    If Not pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        Set pConfig(STR_CONFIGKEY_MODULEPATHS) = New Dictionary
    End If
    Set ModulePathsDictForWriting = pConfig(STR_CONFIGKEY_MODULEPATHS)

End Property


'//
'// Library References configuration
'//

'// The number of references declared in the references configuration collection
Public Property Get ReferencesCount() As Long
    Dim collRefs    As Collection

    If pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Else
        Set collRefs = New Collection
    End If

    ReferencesCount = collRefs.Count
End Property

'// Get the name of a reference declared in the references configuration
Public Property Get ReferenceName(ByVal index As Long) As String
    Dim collRefs    As Collection
    Dim dictRef     As Dictionary

    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Set dictRef = collRefs(index)
    ReferenceName = dictRef(STR_CONFIGKEY_REFERENCE_NAME)
End Property

'// Update the references list with a reference from the VBE.
'// If the reference name is already in the list, that entry is overwritten
'// otherwise a new entry is created.
Public Sub ReferencesUpdateFromVBRef(ByVal VBRef As Reference)
    Dim lngIndex   As Long

    For lngIndex = 1 To ReferencesCount
        If ReferenceName(lngIndex) = VBRef.Name Then
            ReferenceSetFromVBRef lngIndex, VBRef
            Exit Sub ' Assume there is no duplicate entries
        End If
    Next lngIndex

    ReferenceAddFromVBRef VBRef
End Sub

'// Set a reference in the references list to represent the given VB reference.
Private Sub ReferenceSetFromVBRef(ByVal index As Long, ByVal VBRef As Reference)
    Dim collRefs    As Collection
    Dim dictRef     As Dictionary

    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Set dictRef = collRefs(index)

    dictRef(STR_CONFIGKEY_REFERENCE_NAME) = VBRef.Name
    dictRef(STR_CONFIGKEY_REFERENCE_DESCRIPTION) = VBRef.Description
    If VBRef.Type = vbext_rk_TypeLib Then
        dictRef(STR_CONFIGKEY_REFERENCE_GUID) = VBRef.GUID
        dictRef(STR_CONFIGKEY_REFERENCE_MAJOR) = VBRef.Major
        dictRef(STR_CONFIGKEY_REFERENCE_MINOR) = VBRef.Minor
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_PATH) Then _
            dictRef.Remove STR_CONFIGKEY_REFERENCE_PATH
    Else
        dictRef(STR_CONFIGKEY_REFERENCE_PATH) = VBRef.FullPath
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_GUID) Then _
            dictRef.Remove STR_CONFIGKEY_REFERENCE_GUID
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_MAJOR) Then _
            dictRef.Remove STR_CONFIGKEY_REFERENCE_MAJOR
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_MINOR) Then _
            dictRef.Remove STR_CONFIGKEY_REFERENCE_MINOR
    End If

End Sub

'// Add a new reference to the reference list to represent the given reference.
Private Sub ReferenceAddFromVBRef(ByVal VBRef As Reference)
    Dim collRefs    As Collection
    Dim dictRef     As Dictionary

    If Not pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        Set pConfig(STR_CONFIGKEY_REFERENCES) = New Collection
    End If

    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Set dictRef = New Dictionary
    collRefs.Add dictRef

    ReferenceSetFromVBRef collRefs.Count, VBRef
End Sub

'// Remove a reference declaration from the references configuration collection
Public Sub ReferenceRemove(ByVal index As Long)

    Dim collRefs    As Collection
    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    collRefs.Remove index

End Sub

'// Remove the entire References declaration from the configuration
Public Sub ReferencesRemove()
    If pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        pConfig.Remove STR_CONFIGKEY_REFERENCES
    End If
End Sub

'// Add a reference in the references configuration list to the VBE references
'// list. Will delete any existing reference in the VBE references list which
'// has the same name.
Public Sub ReferenceAddToVBRefs(ByVal index As Long, ByVal VBRefs As References)

    Dim collRefs As Collection
    Dim dictRef As Dictionary
    Dim boolLibTypeValid As Boolean
    Dim boolPathTypeValid As Boolean

    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Set dictRef = collRefs(index)

    boolLibTypeValid = _
        dictRef.Exists(STR_CONFIGKEY_REFERENCE_GUID) And _
        dictRef.Exists(STR_CONFIGKEY_REFERENCE_MAJOR) And _
        dictRef.Exists(STR_CONFIGKEY_REFERENCE_MINOR)
    boolPathTypeValid = _
        dictRef.Exists(STR_CONFIGKEY_REFERENCE_PATH)

    If Not (boolLibTypeValid Or boolPathTypeValid) Then
        ' TODO This reference isn't valid. Should at least warn the user?
        Exit Sub
    End If

    If CollectionKeyExists(VBRefs, ReferenceName(index)) Then
        VBRefs.Remove VBRefs(ReferenceName(index))
    End If

    If boolLibTypeValid Then
        VBRefs.AddFromGuid _
            GUID:=dictRef(STR_CONFIGKEY_REFERENCE_GUID), _
            Major:=dictRef(STR_CONFIGKEY_REFERENCE_MAJOR), _
            Minor:=dictRef(STR_CONFIGKEY_REFERENCE_MINOR)
    Else 'PathTypeValid
        VBRefs.AddFromFile _
            Filename:=dictRef(STR_CONFIGKEY_REFERENCE_PATH)
    End If

End Sub

'// Add all references in the references configuration list to the VBE
'// references list. It could be argued that this doesn't logically belong
'// in this class, but it's here because it's fairly trivial addition to
'// ReferenceAddToVBRefs.
Public Sub ReferencesAddToVBRefs(ByVal VBRefs As References)
    Dim lngIndex As Long
    For lngIndex = 1 To ReferencesCount
        ReferenceAddToVBRefs lngIndex, VBRefs
    Next lngIndex
End Sub


'//
'// Serialize and deserialize the configuration from the configuration file
'//

'// Reads and deserializes configuration from the configuration file for the
'// current project.
Public Sub ReadFromProjectConfigFile()

    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigJson       As String

    '// Read JSON file
    If FSO.FileExists(ProjectConfigPath) Then
        Set tsConfigStream = FSO.OpenTextFile(ProjectConfigPath, ForReading)
        strConfigJson = tsConfigStream.ReadAll()
        tsConfigStream.Close
        Set pConfig = JsonConverter.ParseJson(strConfigJson)
    Else
        Set pConfig = New Dictionary
    End If

End Sub


'// Serializes and writes configuration to the configuration file for the
'// current project.
Public Sub WriteToProjectConfigFile()

    Dim tsConfigStream      As Scripting.TextStream
    Dim strConfigJson       As String

    '// Write JSON to file
    strConfigJson = JsonConverter.ConvertToJson(pConfig, vbTab)
    strConfigJson = strConfigJson & vbNewLine
    Set tsConfigStream = FSO.CreateTextFile(ProjectConfigPath, True)
    tsConfigStream.Write strConfigJson
    tsConfigStream.Close

End Sub


'//
'// File system paths for the current context (Project)
'//

'// The directory containing the project file for me.Project
Private Property Get ProjectDir() As String
    ProjectDir = FSO.GetParentFolderName(Project.Filename)
End Property

'// The path to the configuration file for the me.Project project
Private Property Get ProjectConfigPath() As String
    ProjectConfigPath = FSO.BuildPath(ProjectDir, STRCONFIGFILENAME)
End Property
