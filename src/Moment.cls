VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Moment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDate As Double


Public Property Let Moment(NewDate As Double)
    pDate = NewDate
End Property

Public Function ToISOString() As String
    ToISOString = Format(pDate, "yyyy-mm-dd hh:nn:ss")
End Function
