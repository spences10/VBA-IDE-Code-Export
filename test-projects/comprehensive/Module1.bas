Attribute VB_Name = "Module1"
Option Explicit

Public Sub Button1_Click()
    Dim textToShow As Class1
    
    Set textToShow = New Class1
    textToShow.Message = _
        "Seems like it worked!" & vbNewLine & _
        "The 7th fib number: " & Sheet1.FibNumber(7) & vbNewLine & _
        "Some random number: " & ThisWorkbook.PseudoRandomNumber(DateTime.Now)
    
    UserForm1.SetLabelMessage textToShow
    UserForm1.Show
End Sub
