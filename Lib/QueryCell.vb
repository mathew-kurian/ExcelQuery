Private c As range

Public Function constructor(cc As range) As QueryCell
    Set c = cc
    Set constructor = Me
End Function

Public Function Update(val) As QueryCell
    c.Value = val
    Set Update = Me
End Function

Public Function Value() As String
    Value = CStr(c.Value)
End Function

Public Function Append(val) As QueryCell
    c.Value = CStr(c.Value) & CStr(val)
    Set Append = Me
End Function

Public Function Bold(Optional b As Boolean = True) As QueryCell
    c.Font.Bold = b
    Set Bold = Me
End Function

Public Function Underline(Optional b As Boolean = True) As QueryCell
    c.Font.Underline = b
    Set BackgroundRed = Me
End Function

Public Function Background(col As Long) As QueryCell
    c.Interior.Color = col
    Set Background = Me
End Function

Public Function Foreground(col As Long) As QueryCell
    c.Font.Color = col
    Set Foreground = Me
End Function

Public Function Right(Optional by As Integer = 1) As QueryCell
    Dim jc: Set jc = New QueryCell
    Set Right = jc.constructor(c.Offset(0, by))
End Function

Public Function Top(Optional by As Integer = -1) As QueryCell
    Dim jc: Set jc = New QueryCell
    Set Top = jc.constructor(c.Offset(by, 0))
End Function

Public Function Bottom(Optional by As Integer = 1) As QueryCell
    Dim jc: Set jc = New QueryCell
    Set Bottom = jc.constructor(c.Offset(by, 0))
End Function

Public Function Left(Optional by As Integer = -1) As QueryCell
    Dim jc: Set jc = New QueryCell
    Set Left = jc.constructor(c.Offset(0, by))
End Function

Public Function SelectRight(by As Integer, Optional self As Boolean = False) As QueryRange
    Dim jc As QueryRange
    Set jc = New QueryRange
    Dim r: Set r = c.Offset(0, 1)
    Dim cols As Object
    Set cols = CreateObject("System.Collections.ArrayList")
    If (self) Then
        cols.Add c
    End If
    For i = 1 To by
        cols.Add r
        Set r = r.Offset(0, 1)
    Next
    Set SelectRight = jc.constructor(cols)
End Function


Public Function SelectBottom(by As Integer, Optional self As Boolean = False) As QueryRange
    Dim jc As QueryRange
    Set jc = New QueryRange
    Dim r: Set r = c.Offset(1, 0)
    Dim cols As Object
    Set cols = CreateObject("System.Collections.ArrayList")
    If (self) Then
        cols.Add c
    End If
    For i = 1 To by
        cols.Add r
        Set r = r.Offset(1, 0)
    Next
    Set SelectBottom = jc.constructor(cols)
End Function

Public Function Column() As Integer
    Column = c.Column - 1
End Function

Public Function Row() As Integer
    Row = c.Row - 1
End Function
