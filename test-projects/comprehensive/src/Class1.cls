VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Hmm... seems like a perfectly reasonable way to store
' a string. The point is to require a reference for the
' program to work. The dictionary object is made available
' via the scripting library.

Private Const KEY_MESSAGE       As String = "Message"
Private dictMessageStorage      As Dictionary

Public Property Get Message() As String
    Message = dictMessageStorage(KEY_MESSAGE)
End Property

Public Property Let Message(ByVal newMessage As String)
    dictMessageStorage(KEY_MESSAGE) = newMessage
End Property

Private Sub Class_Initialize()
    Set dictMessageStorage = New Dictionary
    dictMessageStorage(KEY_MESSAGE) = vbNullString
End Sub
