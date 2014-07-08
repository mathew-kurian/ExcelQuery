Private cols As Object

Public Function constructor(r As Object) As QueryRange
    Set cols = r
    Set constructor = Me
End Function

Public Function Contains(text As String) As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        Dim pos As Integer
        pos = InStr(CStr(cell.Value), text)
        If (pos > 0) Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set Contains = qr.constructor(List)
End Function

Public Function Exactly(text As String) As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        If (text = "") Then
            If (IsEmpty(cell.Value)) Then
            List.Add cell
            End If
        ElseIf (StrComp(CStr(cell.Value) = text, vbTextCompare) = 0) Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set Exactly = qr.constructor(List)
End Function

Public Function GreaterThan(val As Double) As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        If (cell.Value > val) Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set GreaterThan = qr.constructor(List)
End Function


Public Function LessThan(val As Double) As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        If (cell.Value < val) Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set LessThan = qr.constructor(List)
End Function


Public Function EqualTo(val As Double) As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        If (cell.Value = val) Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set EqualTo = qr.constructor(List)
End Function

Public Function NotEmpty() As QueryRange
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")

    For Each cell In cols
        If (Not cell.Value = "") Then
        List.Add cell
        End If
    Next

    Dim qr As QueryRange
    Set qr = New QueryRange
    Set NotEmpty = qr.constructor(List)
End Function

Public Function Length() As Integer
    Length = cols.Count
End Function

Public Function Eq(Index As Integer) As QueryCell
    Dim jc As QueryCell
    Set jc = New QueryCell
    Set Eq = jc.constructor(cols.Item(Index))
End Function


Public Function Last() As QueryCell
    Dim jc As QueryCell
    Set jc = New QueryCell
    Set Last = jc.constructor(cols.Item(cols.Count - 1))
End Function

Public Function Union() As QueryCell
    Dim jc As QueryCell: Set jc = New QueryCell
    Dim joined As range: Set joined = cols.Item(0)
    For Each cell In cols
        Set joined = range(joined, cell)
    Next
    Set Union = jc.constructor(joined)
End Function

Public Function EqAll() As Object
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")
    For Each cell In cols
        Dim qc: Set qc = New QueryCell
        qc.constructor (cell)
        List.Add (qc)
    Next
    Set EqAll = List
End Function

Public Function Occurs(Optional i As Integer = 1) As Object
    Dim dic: Set dic = CreateObject("Scripting.Dictionary")
    For Each cell In cols
        If (Not dic.Exists(cell.Value)) Then
            dic.Add cell.Value, CreateObject("System.Collections.ArrayList")
        End If
        dic(cell.Value).Add (cell)
    Next
    Dim List As Object
    Set List = CreateObject("System.Collections.ArrayList")
    For Each Key In dic.Keys
        If (dic(Key).Count >= i) Then
            List.Add (dic(Key).Item(0))
        End If
    Next
    Dim jc: Set jc = New QueryRange
    Set Occurs = jc.constructor(List)
End Function

Public Function Update(val As String, Optional reg As String = ",") As QueryRange
    Dim WrdArray() As String
    WrdArray() = Split(val, reg)
    For i = LBound(WrdArray) To UBound(WrdArray)
        If (i < cols.Count) Then
            cols.Item(i).Value = WrdArray(i)
        End If
    Next
    Set Update = Me
End Function
