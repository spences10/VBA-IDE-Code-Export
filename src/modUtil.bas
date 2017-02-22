Attribute VB_Name = "modUtil"
Option Explicit

'// We just need the FSO procedures, no state is necessary
Public FSO As New FileSystemObject


'// Hack to check if Collection key exists
Public Function CollectionKeyExists(ByVal coll As Object, ByVal key As String) As Boolean

    On Error Resume Next
    coll (key)
    CollectionKeyExists = (Err.Number = 0)
    On Error GoTo 0

End Function


'// Display a friendly dialog and return true if user wants to debug
Public Function HandleCrash(ByVal ErrNumber As Long, ByVal ErrDesc As String, ByVal ErrSource As String) As Boolean

    Dim UserAction As Integer

    UserAction = MsgBox( _
        Prompt:= _
            "An unexpected problem occured. Please report this to " & _
            "https://github.com/spences10/VBA-IDE-Code-Export/issues" & vbNewLine & vbNewLine & _
            "Error Number: " & ErrNumber & vbNewLine & _
            "Error Description: " & ErrDesc & vbNewLine & _
            "Error Source: " & ErrSource & vbNewLine & vbNewLine & _
            "Would you like to debug?", _
        Buttons:=vbYesNo + vbDefaultButton2, _
        Title:="Unexpected problem")

    HandleCrash = UserAction = vbYes

End Function
