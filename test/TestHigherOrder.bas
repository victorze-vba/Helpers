Attribute VB_Name = "TestHigherOrder"
Option Explicit

Sub TestForEachAndFilter()
    ForEach Filter(RangeNum(1, 10), "EsPar"), "Imprimir"
End Sub
Function EsPar(number As Long)
    If number Mod 2 = 0 Then
        EsPar = True
    Else
        EsPar = False
    End If
End Function


Sub TestMap()
    ForEach Map(RangeNum(1, 10), "Cuadrado"), "Imprimir"
End Sub
Function Cuadrado(number As Long)
    Cuadrado = number * number
End Function
Sub Imprimir(data As String)
    Debug.Print data
End Sub


Sub TestForEach()
    ForEach RangeNum(1, 10), "PrintSheet"
End Sub
Function PrintSheet(data As String)
    Dim f As Long
    f = Hoja1.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Hoja1.Cells(f, 1) = data
End Function
Sub TestForEachRange()
    Dim rango As Range
    Dim r As Range
    Set rango = Hoja1.Range("A2:A7")
    
    ForEach rango, "Imprimir"
End Sub


Sub TestReduceSuma()
    Debug.Print Reduce(RangeNum(1, 10), "Suma")
    Debug.Print Reduce(RangeNum(1, 10), "Max")
    Debug.Print Reduce(RangeNum(1, 10), "Min")
End Sub
Function Suma(a As Long, b As Long) As Long
    Suma = a + b
End Function
Function Min(a As Long, b As Long) As Long
    Min = IIf(a < b, a, b)
End Function
Function Max(a As Long, b As Long) As Long
    Max = IIf(a > b, a, b)
End Function
Sub TestRangeReduce()
    Dim r As Range
    Set r = Range("A2:A7")
    
    Debug.Print Reduce(r, "Suma")
End Sub


Sub TestSome()
    Debug.Print Some(RangeNum(1, 15), "IsBiggerThan10")
End Sub
Function IsBiggerThan10(number) As Boolean
    IsBiggerThan10 = number > 10
End Function


Sub TestEvery()
    Debug.Print Every(RangeNum(1, 3), "EsPar")
    Debug.Print Some(RangeNum(2, 6, 2), "EsPar")
End Sub


Sub TestFindIndex()
    Debug.Print FindIndex(RangeNum(1, 4), "EsPar")
    Debug.Print FindIndex(RangeNum(1, 7, 2), "EsPar")
End Sub


Function RangeNum(first As Long, last As Long, _
                  Optional step As Variant = Empty) As Collection
    Dim data As New Collection
    Dim i As Long

    If IsEmpty(step) Then
        step = IIf(first < last, 1, -1)
    End If

    For i = first To last Step step
        data.Add i
    Next

    Set RangeNum = data
End Function
