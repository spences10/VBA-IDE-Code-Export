Attribute VB_Name = "modExcelToXML"
Option Explicit

Const ms As Double = 0.000000011574

Sub excelToXML()
    
    If Not changeFileExtension(ThisWorkbook.Path, ThisWorkbook.Name) Then
            
    End If
        
    If Not windowsUnZip(Left(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - 4) & "zip", ThisWorkbook.Path) Then
    
    End If
        
End Sub

Function changeFileExtension(strPath As String, strFileName As String) As Boolean
    
    Dim FSO As New FileSystemObject
    Dim strBaseName As String
    Dim strExtension As String

    On Error GoTo errHandler
    
    If Right(strPath, Len(strPath) - 1) <> Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
    End If
    
    strBaseName = FSO.GetBaseName(strPath & strFileName)
    strExtension = FSO.GetExtensionName(strPath & strFileName)
    
    Call FSO.CopyFile(strPath & strFileName, strPath & strBaseName & ".zip")
    
    changeFileExtension = True

Exit Function

errHandler:
        
    changeFileExtension = False
    
End Function


Function windowsUnZip(strUnzipFullPath, strUnzipDestination) As Boolean
    
    On Error GoTo errHandler
    
    Dim FSO             As FileSystemObject
    Dim fldr            As Folder
    Dim objZipApp       As Shell32.Shell
    
    Dim strFolderPath   As String
    Dim strFolderName   As String
    Dim strZipFolder    As String
    
    Set FSO = New FileSystemObject
    Set objZipApp = New Shell
    
    strFolderPath = FSO.GetParentFolderName(strUnzipFullPath)
    strFolderName = FSO.GetBaseName(strUnzipFullPath)
    
    '// TODO this needs to be a function!!!
    If Right(strFolderPath, Len(strFolderPath) - 1) <> Application.PathSeparator Then
        strFolderPath = strFolderPath & Application.PathSeparator
    End If
    
    strZipFolder = strFolderPath & strFolderName
    
    '// create folder of zip
    If FSO.FolderExists(strZipFolder) = True Then
        '// Do nothing
    Else
        '// or you can create it
        Set fldr = FSO.CreateFolder(strZipFolder)
    
        If fldr Is Nothing Then
            MsgBox "Could not create the folder"
        End If
    End If
    
    With objZipApp
        .Namespace(strFolderPath & strFolderName).Copyhere .Namespace(strUnzipFullPath).Items
    End With
    
    windowsUnZip = True
    
Exit Function

errHandler:
    
    windowsUnZip = False
    
End Function


Function windowsZip(strFilePath, strZipFile)
  
  Dim objZipShell As WshShell
  Dim objZipFSO As FileSystemObject
  Dim objZipApp As Shell32.Shell
  Dim lngZipFileCount As Long
  
  Dim objFileNameInZip As Object
  
  Dim lngLoop As Long
  
  Dim strFilePathArray() As String
  Dim strFileName As String
  
  Dim blnDupe As Boolean
   
  Set objZipShell = New WshShell
  Set objZipFSO = New FileSystemObject
  
  If Not objZipFSO.FileExists(strZipFile) Then
    Call newZip(strZipFile)
  End If

  Set objZipApp = New Shell
  
  lngZipFileCount = objZipApp.Namespace(strZipFile).Items.Count

  strFilePathArray = Split(strFilePath, "\")
  strFileName = (strFilePathArray(UBound(strFilePathArray)))
  
  '// listfiles
  blnDupe = False
  For Each objFileNameInZip In objZipApp.Namespace(strZipFile).Items
    If LCase(strFileName) = LCase(objFileNameInZip) Then
      blnDupe = True
      Exit For
    End If
  Next
  
  If Not blnDupe Then
    objZipApp.Namespace(strZipFile).Copyhere strFilePath

    '// Wait until Compressing is done
    On Error Resume Next
    lngLoop = 0
    Do Until lngZipFileCount < objZipApp.Namespace(strZipFile).Items.Count
      Application.Wait Now + ms * 100
      lngLoop = lngLoop + 1
    Loop
    On Error GoTo 0
  End If

End Function

Sub newZip(strNewZip)
    
    Dim objNewZipFSO As FileSystemObject
    Dim objNewZipFile As TextStream
  
    Set objNewZipFSO = New FileSystemObject
    Set objNewZipFile = objNewZipFSO.CreateTextFile(strNewZip)
      
    objNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
    
    objNewZipFile.Close
    Set objNewZipFSO = Nothing

    Application.Wait Now + ms * 500   '// 500 milliseconds
    
End Sub


