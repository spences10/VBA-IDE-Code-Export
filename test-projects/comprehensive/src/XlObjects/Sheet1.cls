VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' An inefficient evaluation of an element of the Fibonacci sequence
Public Function FibNumber(ByVal NumberIndex As Long)
    If NumberIndex <= 2 Then
        FibNumber = 1
    Else
        FibNumber = FibNumber(NumberIndex - 1) + FibNumber(NumberIndex - 2)
    End If
End Function
