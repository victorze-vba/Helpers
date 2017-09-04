Attribute VB_Name = "TetsMoment"
Option Explicit

Sub test_moment()
    Call test_moment_new_date
    
    Call test_moment_manipulate_add
    
    Call test_moment_manipulate_start_of
    
    Call test_moment_manipulate_end_of
    
    Call test_moment_display
End Sub

Sub test_moment_new_date()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    
    Specs.Description = "Definición de una fecha nueva"
    
    ' Fecha
    Set MyMoment = New Moment
    MyMoment.Moment = CDate(43093)
    
    With Specs.It("crear una fecha con tipo 'Date'")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:00:00"
    End With
    
    ' Fecha con horas minutos y segundos
    Set MyMoment = New Moment
    MyMoment.Moment = CDate(43093.43)
    
    With Specs.It("crear una fecha con tipo 'Date' con horas, minutos y segundos")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:19:12"
    End With

    ' Fecha con numero entero
    Set MyMoment = New Moment
    MyMoment.Moment = 43093
    
    With Specs.It("crear una fecha con tipo 'Integer'")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:00:00"
    End With
    
    ' Fecha con número decimal
    Set MyMoment = New Moment
    MyMoment.Moment = 43093.43
    
    With Specs.It("crear una fecha con tipo 'Double'")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:19:12"
    End With
    
    InlineRunner.RunSuite Specs, True, False, True
End Sub

Sub test_moment_manipulate_add()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    Dim ReferenceMoment As Date
    
    Specs.Description = "Manupulación de momentos 'Add'"
    
    ReferenceMoment = DateSerial(2017, 12, 24) ' 24/12/2017
    
    ' Sumar restar Años
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 3, Years

    With Specs.It("debe sumar 3 años")
        .Expect(MyMoment.ToISOString).ToEqual "2020-12-24 00:00:00"
    End With
    
    MyMoment.Add -5, Years
    
    With Specs.It("debe restar 5 años")
        .Expect(MyMoment.ToISOString).ToEqual "2015-12-24 00:00:00"
    End With
    
    ' Sumar restar meses
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 3, Months
    
    With Specs.It("debe sumar 3 meses")
        .Expect(MyMoment.ToISOString).ToEqual "2018-03-24 00:00:00"
    End With
    
    MyMoment.Add -5, Months

    With Specs.It("debe restar 5 meses")
        .Expect(MyMoment.ToISOString).ToEqual "2017-10-24 00:00:00"
    End With
    
    ' Sumar restar días
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 3, Days
    
    With Specs.It("debe sumar 3 días")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-27 00:00:00"
    End With
    
    MyMoment.Add -5, Days

    With Specs.It("debe restar 5 días")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-22 00:00:00"
    End With
    
    ' Sumar restar horas
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 15, Hours
    
    With Specs.It("debe sumar 15 horas")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 15:00:00"
    End With
    
    MyMoment.Add -10, Hours

    With Specs.It("debe restar 10 horas")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 05:00:00"
    End With
    
    ' Sumar restar minutos
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 45, Minutes
    
    With Specs.It("debe sumar 45 minutos")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:45:00"
    End With
    
    MyMoment.Add -30, Minutes
    
    With Specs.It("debe restar 30 minutos")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:15:00"
    End With

    ' Sumar restar segundos
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.Add 20, Seconds
    
    With Specs.It("debe sumar 20 segundos")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:00:20"
    End With
    
    MyMoment.Add -15, Seconds
    
    With Specs.It("debe restar 15 segundos")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:00:05"
    End With
    
    InlineRunner.RunSuite Specs, True, False, True
End Sub

Sub test_moment_manipulate_start_of()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    Dim ReferenceMoment As Date
    
    Specs.Description = "Manipulación de momentos 'Start Of'"
    
    ReferenceMoment = CDate(43093.43) ' 24/12/2017 10:19:12 a.m.
    
    ' Cambiar a inicio de...
    ' Año
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoStart OfYear
    
    With Specs.It("debe cambiar a inicio del año")
        .Expect(MyMoment.ToISOString).ToEqual "2017-01-01 00:00:00"
    End With
    
    ' Mes
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoStart OfMonth

    With Specs.It("debe cambiar a inicio del mes")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-01 00:00:00"
    End With
    
    ' Día
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoStart OfDay

    With Specs.It("debe cambiar a inicio del día")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 00:00:00"
    End With
    
    ' Hora
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoStart OfHour

    With Specs.It("debe cambiar a inicio de la hora")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:00:00"
    End With
    
    ' Minuto
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoStart OfMinute

    With Specs.It("debe cambiar a inicio del minuto")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:19:00"
    End With

    InlineRunner.RunSuite Specs, True, False, True
End Sub

Sub test_moment_manipulate_end_of()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    Dim ReferenceMoment As Date
    
    Specs.Description = "Manipulación de momentos 'End Of'"
    
    ReferenceMoment = CDate(43093.43) ' 24/12/2017 10:19:12 a.m.
    
    ' Cambiar a fin de...
    ' Año
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoEnd OfYear
    
    With Specs.It("debe cambiar a fin de año")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-31 23:59:59"
    End With
    
    ' Mes
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoEnd OfMonth
    
    With Specs.It("debe cambiar a fin de mes")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-31 23:59:59"
    End With
    
    ' Día
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoEnd OfDay
    
    With Specs.It("debe cambiar a fin del día")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 23:59:59"
    End With
    
    ' Hora
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoEnd OfHour
    
    With Specs.It("debe cambiar al final de la hora")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:59:59"
    End With
    
    ' Minuto
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    MyMoment.GoEnd OfMinute
    
    With Specs.It("debe cambiar al final del minuto")
        .Expect(MyMoment.ToISOString).ToEqual "2017-12-24 10:19:59"
    End With
    
    InlineRunner.RunSuite Specs, True, False, True
End Sub

Sub test_moment_display()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    Dim RefMoment As New Moment
    Dim ReferenceMoment As Date
    
    Specs.Description = "Mostrar el momento según formato"
    
    ReferenceMoment = CDate(43093.43) ' 24/12/2017 10:19:12 a.m.
    
    Set MyMoment = New Moment
    MyMoment.Moment = ReferenceMoment
    
    With Specs.It("debe mostrar el momento formato 'dd/mm/yyyy'")
        .Expect(MyMoment.ToFormat()).ToEqual "24/12/2017"
    End With
    
    With Specs.It("debe mostrar el momento en formato 'dd - mm - yyyy'")
        .Expect(MyMoment.ToFormat("dd - mm - yyyy")).ToEqual "24 - 12 - 2017"
    End With
    
    With Specs.It("debe mostrar el momento texto, ej: domingo, 24 de diciembre de 2017")
        .Expect(MyMoment.ToString()).ToEqual "domingo, 24 de diciembre de 2017"
    End With
    
    With Specs.It("debe mostrar el momento en valor")
        .Expect(MyMoment.ValueOf()).ToEqual 43093.43
    End With
    
    ' diff
    RefMoment.Moment = 43510.3  ' 14/02/2019 07:12:00 a.m.
    
    With Specs.It("debe mostrar la diferencia entre dos momentos en días")
        .Expect(MyMoment.DiffDays(RefMoment)).ToEqual 417
    End With
    
    ' get
    With Specs.It("debe mostrar el año del momento")
        .Expect(MyMoment.GetYear).ToEqual 2017
    End With
    
    With Specs.It("debe mostrar el mes del momento")
        .Expect(MyMoment.GetMonth).ToEqual 12
    End With
    
    With Specs.It("debe mostrar el día del momento")
        .Expect(MyMoment.GetDay).ToEqual 24
    End With
    
    InlineRunner.RunSuite Specs, True, False, True
End Sub
