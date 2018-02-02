Attribute VB_Name = "HigherOrder"
Option Explicit

Function Filter(ByVal iterable As Variant, callback As String) As Collection
    Dim passed As New Collection
    Dim element As Variant
    
    For Each element In iterable
        If (Application.Run(callback, element)) Then
            passed.Add element
        End If
    Next element
    
    Set Filter = passed
End Function

Function Map(ByVal iterable As Variant, callback As String) As Collection
    Dim mapped As New Collection
    Dim element As Variant
    
    For Each element In iterable
        mapped.Add Application.Run(callback, element)
    Next element
    
    Set Map = mapped
End Function

Sub ForEach(ByVal iterable As Variant, callback As String)
    Dim element As Variant
    
    For Each element In iterable
        Application.Run callback, element
    Next element
End Sub

Function Reduce(ByVal iterable As Variant, callback As String, _
                Optional start As Variant = Empty) As Variant
    Dim current As Variant
    Dim element As Variant
    
    If IsEmpty(start) Then
        start = iterable.Item(1)
    End If
    
    current = start
    For Each element In iterable
        current = Application.Run(callback, current, element)
    Next element
    
    Reduce = current
End Function

Function Some(ByVal iterable As Variant, callback As String) As Boolean
    Dim element As Variant
    
    For Each element In iterable
        If Application.Run(callback, element) Then
            Some = True
            Exit Function
        End If
    Next element
    
    Some = False
End Function

Function Every(ByVal iterable As Variant, callback As String) As Boolean
    Dim element As Variant
    
    For Each element In iterable
        If Not Application.Run(callback, element) Then
            Every = False
            Exit Function
        End If
    Next element
    
    Every = True
End Function

Function FindIndex(ByVal iterable As Variant, callback As String) As Long
    Dim i As Long
    
    For i = 1 To iterable.Count
        If Application.Run(callback, iterable.Item(i)) Then
            FindIndex = i
            Exit Function
        End If
    Next i
    
    FindIndex = -1
End Function
