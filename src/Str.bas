Attribute VB_Name = "Str"
''
' Scraping v0.1.1 Alpha
' (c) Victor Zevallos - https://github.com/vba-dev/vba-scraping
'
' @Module Str
' @author victorzevallos@protonmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Public Function Upper(Str As String) As String
    Upper = UCase(Str)
End Function

Public Function Lower(Str As String) As String
    Lower = LCase(Str)
End Function

Public Function Count(Str As String) As Integer
    Count = Len(Str)
End Function

Public Function Reverse(Str As String) As String
    Reverse = StrReverse(Str)
End Function

Public Function Find(Str As String, SubStr As String, Optional Start As Integer = 1) As Integer
    Find = InStr(Start, Str, SubStr)
End Function

Public Function Replaces(Find As String, Replace_ As String, Str As String) As String
    Replaces = Replace(Str, Find, Replace_)
End Function

Public Function Slice(Str As String, StartIndex As Integer, Optional EndIndex As Integer = 1000) As String
    Dim Length As Integer
    
    Length = EndIndex - StartIndex + 1
    
    Slice = Mid(Str, StartIndex, Length)
End Function

Public Function Strip(Str As String) As String
    Strip = Trim(Str)
End Function

Public Function Capitalice(Str As String) As String
    Str = LCase(Trim(Str))
    
    Capitalice = UCase(Left(Str, 1)) & Right(Str, Len(Str) - 1)
End Function

Public Function Title(ByVal Str As String) As String
    Dim i As Integer
    Dim Char As String
    Dim part1 As String
    Dim part2 As String
    Dim NewStr As String
    Dim PrevChar As String
    
    Str = Capitalice(Trim(LCase(Str)))
    NewStr = Str
    For i = 1 To Len(Str)
        Char = Mid(Str, i, 1)
        
        If PrevChar = " " And Char <> " " Then
            part1 = Mid(NewStr, 1, i - 1)
            part2 = Capitalice(Mid(NewStr, i))
            NewStr = part1 & part2
        End If
        
        PrevChar = Char
    Next i
    
    Title = NewStr
End Function

Public Function Words(Str As String) As Integer
    Dim i As Integer
    Dim Char As String
    Dim Count As Integer
    Dim Previous As String
    
    Str = Trim(Str)
    Previous = " "
    
    For i = 1 To Len(Str)
        Char = Mid(Str, i, 1)
        
        If Previous = " " And Char <> " " Then
            Count = Count + 1
        End If
        
        Previous = Char
    Next i
    
    Words = Count
End Function

Public Function Char(Str As String, Index As Integer) As String
    Char = Mid(Trim(Str), Index, 1)
End Function

