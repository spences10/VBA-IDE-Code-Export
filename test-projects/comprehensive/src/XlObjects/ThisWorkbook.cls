VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' A made up PRNG (probably not a good one).
Public Function PseudoRandomNumber(ByVal Seed As Double) As Double
    Seed = Seed - Int(Seed)
    Seed = 234945.3452 * Seed + 6352.5435
    Seed = Seed - Int(Seed)
    PseudoRandomNumber = Seed
End Function
