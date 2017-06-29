Attribute VB_Name = "Str_Test"
Option Explicit

Sub StrTest()
    Dim Specs As New SpecSuite

    With Specs.It("Upper")
        .Expect(Str.Upper("excel")).ToEqual "EXCEL"
    End With
    
    With Specs.It("Lower")
        .Expect(Str.Lower("MACROS")).ToEqual "macros"
    End With
    
    With Specs.It("Count")
        .Expect(Str.Count("VBA Excel")).ToEqual 9
    End With

    With Specs.It("Reverse")
        .Expect(Str.Reverse("EXCEL")).ToEqual "LECXE"
    End With

    With Specs.It("Find")
        .Expect(Str.Find("VBA reference for Excel", "ce")).ToEqual 12
        .Expect(Str.Find("VBA reference for Excel", "ce", 13)).ToEqual 21
    End With

    With Specs.It("Replace")
        .Expect(Str.Replaces("Excel", "Access", "VBA reference for Excel")).ToEqual "VBA reference for Access"
        .Expect(Str.Replaces("Word", "Access", "VBA reference for Excel")).ToEqual "VBA reference for Excel"
    End With

    With Specs.It("Slice")
        .Expect(Str.Slice("VBA reference for Excel", 5)).ToEqual "reference for Excel"
        .Expect(Str.Slice("VBA reference for Excel", 15, 17)).ToEqual "for"
        .Expect(Str.Slice("", 15)).ToEqual ""
    End With

    With Specs.It("Strip")
        .Expect(Str.Strip("  VBA ")).ToEqual "VBA"
    End With
    
    With Specs.It("Capitalice")
        .Expect(Str.Capitalice("JOHN LENnon")).ToEqual "John lennon"
        .Expect(Str.Capitalice("w.")).ToEqual "W."
    End With

    With Specs.It("Title")
        .Expect(Str.Title("JOHN LENnon")).ToEqual "John Lennon"
        .Expect(Str.Title("   john w. leNNOn x.")).ToEqual "John W. Lennon X."
    End With

    With Specs.It("Words")
        .Expect(Str.Words("VBA Excel")).ToEqual 2
        .Expect(Str.Words("")).ToEqual 0
        .Expect(Str.Words("  Excel   Word    Access   ")).ToEqual 3
    End With

    With Specs.It("Char")
        .Expect(Str.Char("VBA Excel", 5)).ToEqual "E"
        .Expect(Str.Char("  VBA Excel", 5)).ToEqual "E"
        .Expect(Str.Char("VBA Excel", 16)).ToEqual ""
    End With
    

    InlineRunner.RunSuite Specs
End Sub

