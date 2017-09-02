Attribute VB_Name = "TetsMoment"
Option Explicit

Sub test_moment()
    Call test_moment_parse
End Sub

Sub test_moment_new_date()
    Dim Specs As New SpecSuite
    Dim MyDate As Moment
    
    Specs.Description = "Definición de una fecha nueva"
    
    Set MyDate = New Moment
    MyDate.Moment = CDate(43093)
    
    With Specs.It("debe crear una fecha con tipo 'Date'")
        .Expect(MyDate.ToISOString).ToEqual "2017-12-24 00:00:00"
    End With
    
    
    Set MyDate = New Moment
    MyDate.Moment = CDate(43093.43)
    
    With Specs.It("debe crear una fecha con tipo 'Date' con horas, minutos y segundos")
        .Expect(MyDate.ToISOString).ToEqual "2017-12-24 10:19:12"
    End With
    
    
    Set MyDate = New Moment
    MyDate.Moment = 43093
    
    With Specs.It("debe crear una fecha con tipo 'Integer'")
        .Expect(MyDate.ToISOString).ToEqual "2017-12-24 00:00:00"
    End With
    
    
    Set MyDate = New Moment
    MyDate.Moment = 43093.43
    
    With Specs.It("debe crear una fecha con tipo 'Double'")
        .Expect(MyDate.ToISOString).ToEqual "2017-12-24 10:19:12"
    End With
    
    InlineRunner.RunSuite Specs, True, False, True
End Sub

Sub test_moment_get_set()

End Sub

Sub test_moment_manipulate()

End Sub

Sub test_moment_display()

End Sub

Sub test_moment_query()

End Sub

Sub test_moment_customize()

End Sub
