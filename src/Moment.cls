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
    If NewMoment <= 0 Then
        Err.Raise 65535, , "El valor del objeto 'Moment' debe ser positivo."
    ElseIf pMoment > 0 Then
        Err.Raise 65534, , "El objeto 'Moment' ya posee un valor."
    Else
        pMoment = NewMoment
    End If
End Property

' --------------------------------------------- '
' Manipular
' --------------------------------------------- '

' Una vez que tenga un 'momento', tal vez quiera manipularlo de alguna manera.

''
' Suma o resta una unidad de tiempo
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

''
' Cambia el 'momento' al inicio de una unidad de tiempo.
'
' @param  MomentTypeOf TimeType (tipo de tiempo a cambiar)
' @return void
''
Public Sub GoStart(TimeType As MomentTypeOf)
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

''
' Cambia el 'momento' al final de una unidad de tiempo.
'
' @param  MomentTypeOf TimeType (tipo de tiempo a cambiar)
' @return void
''
Public Sub GoEnd(TimeType As MomentTypeOf)
    Select Case TimeType
        Case OfYear
            pMoment = DateSerial(Year(pMoment) + 1, 1, 1)
            Add -1, Seconds
        Case OfMonth
            pMoment = DateSerial(Year(pMoment), Month(pMoment) + 1, 1)
            Add -1, Seconds
        Case OfDay
            pMoment = DateSerial(Year(pMoment), Month(pMoment), Day(pMoment) + 1)
            Add -1, Seconds
        Case OfHour
            Add 59 - Minute(pMoment), Minutes
            Add 59 - Second(pMoment), Seconds
        Case Minutes
            Add 59 - Second(pMoment), Seconds
    End Select
End Sub


' --------------------------------------------- '
' Mostrar
' --------------------------------------------- '

' Despúes de hacer una manipulación, usted necesita alguna forma de mostrar el 'momento'

''
' Retorna el 'momento' según formato.
'
' @param  String FormatDate (formato de fecha, por defecto: "dd/mm/yyyy")
' @return String
''
Public Function ToFormat(Optional FormatDate As String = "dd/mm/yyyy") As String
    ToFormat = Format(pMoment, FormatDate)
End Function

''
' Retorna la 'momento' en formato iso "yyyy-mm-dd hh:nn:ss"
'
' @return String
''
Public Function ToISOString() As String
    ToISOString = Format(pMoment, "yyyy-mm-dd hh:nn:ss")
End Function

''
' Retorna el 'momento' en texto, ej: domingo, 24 de diciembre de 2017
'
' @return String
''
Public Function ToString()
    ToString = FormatDateTime(pMoment, vbLongDate)
End Function

''
' Retorna el momento en formato 'Double'
'
' @return Double
''
Public Function ValueOf() As Double
    ValueOf = pMoment
End Function

''
' Retorna la diferencia en días entre dos momentos
'
' @param  Moment RefDate (fecha de referencia)
' @return Double
''
Public Function DiffDays(RefDate As Moment) As Double
    DiffDays = DateDiff("d", pMoment, RefDate.ValueOf)
End Function

''
' Retorna el año del momento. ej 2018
'
' @return Integer
''
Public Function GetYear() As Integer
    GetYear = Year(pMoment)
End Function

''
' Retorna el mes del momento. 1-12
'
' @return Integer
''
Public Function GetMonth() As Integer
    GetMonth = Month(pMoment)
End Function

''
' Retorna el día del momento. 1-31
'
' @return Integer
''
Public Function GetDay() As Integer
    GetDay = Day(pMoment)
End Function
