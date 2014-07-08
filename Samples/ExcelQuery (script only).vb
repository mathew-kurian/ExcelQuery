Sub ExcelQuery()

' Required Variables
Dim Excel: Set Excel = New Query
Dim Color: Set Color = New QueryColor

' Delete these worksheets if they exist
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Delete
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers - Output").Delete

' Making a table with query
' This is to make the table using VBA
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).SelectRight(5, True).Update("Field,FullName,PayCode,Salary,Bonus").Union.Bold
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).SelectBottom(6).Update ("Engineering,Medicine,Travel,Musician,Medicine,Medical,Animals,Dogs")
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).Right.SelectBottom(6).Update ("Sam,Tyrion,Joffrey,Paul,Mountain,Gregor")
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).Right.Right.SelectBottom(6).Update("40,45,10,25,45,10").Union.Bold
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).Right.Right.Right.SelectBottom(6).Update ("343654,456454,1000000,3456,858454,343645")
Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).Right.Right.Right.Right.SelectBottom(6).Update ("10,3,4,1,42,3")

' Searching with query. Find people in the medicine background and then find their total salary
' This searches for each field and then moves it over to a new sheet
For Each field In Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).SelectBottom(10).NotEmpty.Occurs(1).EqAll
    Dim totalSalary: totalSalary = 0
    For Each cell In Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).SelectBottom(10).Contains(field.Value()).EqAll
        totalSalary = totalSalary + cell.Right(Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers").Cells("A1").Eq(0).SelectRight(10).Contains("Salary").Eq(0).Column).Background(Color.RED).Foreground(Color.WHITE).Value
    Next
    Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers - Output").Cells("A1").Eq(0).SelectRight(10, True).Exactly("").Eq(0).Update (field.Value)
    Excel.Workbook(ThisWorkbook).Worksheet("VBA Headers - Output").Cells("A1").Eq(0).Bottom.SelectRight(10, True).Exactly("").Eq(0).Update (totalSalary)
Next

End Sub
