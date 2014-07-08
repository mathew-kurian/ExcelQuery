Private wbook As Workbook

Public Function constructor(book As Workbook) As QueryBook
    Set wbook = Nothing
    If (IsNull(book)) Then
        Set wbook = ActiveWorkbook
    Else
        Set wbook = book
    End If
    Set constructor = Me
End Function

Public Function Name() As String
    Dim lastError As String
    On Error GoTo ErrorHandler
    lastError = "Workbook is nothing"
    Name = wbook.Name
    Exit Function
ErrorHandler:
    MsgBox lastError
    End
End Function

Public Function Worksheet(Name As String) As QuerySheet
    Dim lastError As String
    On Error GoTo CreateWorksheet
    Dim qsheet As QuerySheet
    Set qsheet = New QuerySheet
    Set Worksheet = qsheet.constructor(wbook.Worksheets(Name))
    Exit Function
CreateWorksheet:
    On Error GoTo ErrorHandler
    lastError = "Cannot add " & Name
    wbook.Worksheets().Add().Name = Name
    Set Worksheet = qsheet.constructor(wbook.Worksheets(Name))
    Exit Function
ErrorHandler:
    MsgBox lastError
    End
End Function
