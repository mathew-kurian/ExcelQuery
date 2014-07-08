Public Function Workbook(wbk As Workbook) As QueryBook
    Dim qb As QueryBook
    Set qb = New QueryBook
    Set Workbook = qb.constructor(wbk)
End Function

