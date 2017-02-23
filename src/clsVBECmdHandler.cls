VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBECmdHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// Add references for :
'//     Microsoft Visual Basic For Applications Extensibility 5.3

Public WithEvents EvtHandler As VBIDE.CommandBarEvents
Attribute EvtHandler.VB_VarHelpID = -1


Private Sub EvtHandler_Click(ByVal CommandBarControl As Object, ByRef Handled As Boolean, ByRef CancelDefault As Boolean)

    On Error Resume Next

    '// Run the code specified in the object's OnAction property.
    Application.Run CommandBarControl.OnAction

    '// Indicate to the Events object that we've successfully handled the event.
    Handled = True
    CancelDefault = True

End Sub
