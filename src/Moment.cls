VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Moment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Moment v0.1.1 Alpha
' (c) Victor Zevallos - https://github.com/vba-dev/vba-scraping
'
' @class Moment: librería para trabajar con datos de tipo 'Date' (fecha)
'
' @author victorzevallos@protonmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Private pMoment As Double

Public Enum MomentType
    Years
    Months
    Days
    Hours
    Minutes
    Seconds
End Enum

Public Enum MomentTypeOf
    OfYear
    OfMonth
    OfDay
    OfHour
    OfMinute
End Enum

''
' Establecer el valor de la fecha (momento)
'
' @param Double NewMoment(fecha de tipo 'Integer', 'Double' o 'Date')
''
Public Property Let Moment(NewMoment As Double)
    pMoment = NewMoment
End Property

Public Function ToISOString() As String
    ToISOString = Format(pMoment, "yyyy-mm-dd hh:nn:ss")
End Function

' --------------------------------------------- '
' Manipular
' --------------------------------------------- '

' Una vez que tenga un 'Moment', tal vez quiera manipularlo de alguna manera.

''
' Sumar o restar una unidad de tiempo
'
' @param  Integer    Number  (cantidad de tiempo que se suma o se resta)
' @param  MomentType AddType (tipo de unidad ej: Days, Hours...)
' @return void
''
Public Sub Add(Number As Integer, TimeType As MomentType)
    Select Case TimeType
        Case Years
            pMoment = DateAdd("YYYY", Number, pMoment)
        Case Months
            pMoment = DateAdd("m", Number, pMoment)
        Case Days
            pMoment = DateAdd("d", Number, pMoment)
        Case Hours
            pMoment = DateAdd("h", Number, pMoment)
        Case Minutes
            pMoment = DateAdd("n", Number, pMoment)
        Case Seconds
            pMoment = DateAdd("s", Number, pMoment)
    End Select
End Sub

Public Sub StartTime(TimeType As MomentTypeOf)
    Select Case TimeType
        Case OfYear
            pMoment = DateSerial(Year(pMoment), 1, 1)
        Case OfMonth
            pMoment = DateSerial(Year(pMoment), Month(pMoment), 1)
        Case OfDay
            pMoment = DateSerial(Year(pMoment), Month(pMoment), Day(pMoment))
        Case OfHour
            Add -Minute(pMoment), Minutes
            Add -Second(pMoment), Seconds
        Case Minutes
            Add -Second(pMoment), Seconds
    End Select
End Sub
