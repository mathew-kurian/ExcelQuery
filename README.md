#ExcelQuery
ExcelQuery is a small, fast library that allows for syntactically understandable cell traversal.

#API
Check out the api [here](https://github.com/bluejamesbond/ExcelQuery/wiki/Home)

#Getting Started
**1**. Open a blank excel workbook

**2**. View Macros

![GitHub Logo](http://i.imgur.com/CHKIr2G.png)

**3**. Set Macro name

![GitHub Logo](http://i.imgur.com/I1J5QPV.png)

**4**. Import each of the `Lib` files
```
     - Query.cls
     - QueryBook.cls
     - QueryCell.cls
     - QueryColor.cls
     - QueryRange.cls
     - QuerySheet.cls
```

![GitHub Logo](http://i.imgur.com/Yyd226a.png?1)

![GitHub Logo](http://i.imgur.com/qTUqjeJ.png?1)

**5**. Double-click `Module 1` and paste in the following:

![GitHub Logo](http://i.imgur.com/5n9Howm.png)
```
Sub InsertARandonName()

' Required Variables
Q.Workbook(ThisWorkbook) _
     .Worksheet("Sheet1") _
     .Cells("A1") _
     .Eq(0) _
     .SelectRight(5, True) _
     .Update("Field,FullName,PayCode,Salary,Bonus") _
     .Union _
     .Bold

End Sub
```
**6**. Make sure the Macro Window looks similar to the following image. Then click the Play button to run the program.

![GitHub Logo](http://i.imgur.com/j1kzBL6.png)

**7**. Go to Sheet1 on the Excel workbook to see the output
