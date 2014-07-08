Private wsheet As Worksheet

Public Function constructor(Worksheet As Worksheet) As QuerySheet
    Set wsheet = Worksheet
    Set constructor = Me
End Function

Public Function Rename(Name As String) As QuerySheet
    wsheet.Name = Name
    Set Rename = Me
End Function

Public Function Name() As String
    Name = wsheet.Name
End Function

Public Function Delete()
    Dim lastError As String
    On Error GoTo ErrorHandler
    lastError = "Cannot delete the last visible worsheet"
    Application.DisplayAlerts = False
    wsheet.Delete
    Application.DisplayAlerts = True
    Exit Function
ErrorHandler:
    MsgBox lastError & ". Press OK to continue execution."
End Function

Public Function Cells(range As String) As QueryRange
    Dim qr As QueryRange
    Dim cols As Object
    Set qr = New QueryRange
    Set cols = CreateObject("System.Collections.ArrayList")
    For Each c In wsheet.range(range).Cells
        cols.Add c
    Next
    Set Cells = qr.constructor(cols)
End Function
