Attribute VB_Name = "TetsMoment"
Option Explicit

Sub test_moment()
    Call test_moment_new_date
    
    Call test_moment_manipulate
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

Sub test_moment_manipulate()
    Dim Specs As New SpecSuite
    Dim MyMoment As Moment
    
    Specs.Description = "Manupulación de momentos"
    
    ' Sumar restar Años
    Set MyMoment = New Moment
    MyMoment.Moment = DateSerial(2017, 12, 24)
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
    MyMoment.Moment = DateSerial(2017, 12, 24)
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
    MyMoment.Moment = DateSerial(2017, 12, 24)
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
    MyMoment.Moment = DateSerial(2017, 12, 24)
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
    MyMoment.Moment = DateSerial(2017, 12, 24)
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
    MyMoment.Moment = DateSerial(2017, 12, 24)
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

Sub test_moment_display()

End Sub

Sub test_moment_query()

End Sub

Sub test_moment_customize()

End Sub
