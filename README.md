#ExcelQuery
ExcelQuery is a small, fast library that allows for syntactically understandable cell traversal.

#API

##Query

| Macro | Returns |
|:---------------|:--------------:|
|`Workbook(Workbook workbook)` | [`QueryBook`](#querybook) |
|`Color()` | [`QueryColor`](#querybook) |
|`Interact()` | [`QueryInteract`](#queryinteract) |

## QueryBook

| Macro | Returns |
|:---------------|:--------------:|
|`Name()` | `String` |
|`WorkSheet(String name)` | [`QuerySheet`](#querysheet) |

## QuerySheet

| Macro | Returns |
|:---------------|:--------------:|
|`Rename(String name)` | [`QuerySheet`](#querysheet) |
|`Name()` | `String` |
|`Delete()` | `NULL` |
|`Cells(String range)` | [`QueryRange`](#queryrange) |

##QueryRange

| Macro | Returns |
|:---------------|:--------------:|
|`Bold(Optional Boolean enable)` | [`QueryRange`](#queryrange) |
|`Underline(Optional Boolean enable)` | [`QueryRange`](#queryrange)  |
|`Background(Long color)` | [`QueryRange`](#queryrange) |
|`Foreground(Long color)` | [`QueryRange`](#queryrange) |
|`Contains(Text)` | [`QueryRange`](#queryrange) |
|`ContainsExactly(Text)` | [`QueryRange`](#queryrange) |
|`ContainsNotExactly(Text)` | [`QueryRange`](#queryrange) |
|`LessThan(Double value)` | [`QueryRange`](#queryrange) |
|`EqualTo(Double value)` | [`QueryRange`](#queryrange) |
|`NotEmpty()` | [`QueryRange`](#queryrange) |
|`Length()` | `Integer` |
|`Eq(Integer index)` | [`QueryCell`](#querycell) |
|`Last()` | [`QueryCell`](#querycell) |
|`Union()` | [`QueryCell`](#querycell) |
|`EqAll()` | `List` |
|`Occurs(Integer integer)` | [`QueryRange`](#queryrange) |
|`Header(Auto value, Optional Auto delimeter)` | [`QueryRange`](#queryrange)
|`Update(Auto value)` | [`QueryRange`](#queryrange)

##QueryCell

| Macro | Returns |
|:---------------|:--------------:|
|`Update(Auto value)` | [`QueryCell`](#querycell) |
|`Value()` | `Double` |
|`Text()` | `String` |
|`Append(Auto value)` | [`QueryCell`](#querycell) |
|`Bold(Optional Boolean enable)` | [`QueryCell`](#querycell) |
|`Underline(Optional Boolean enable)` | [`QueryCell`](#querycell)  |
|`Background(Long color)` | [`QueryCell`](#querycell) |
|`Foreground(Long color)` | [`QueryCell`](#querycell) |
|`Right(Optional Integer count)` | [`QueryCell`](#querycell) |
|`Up(Optional Integer count)` | [`QueryCell`](#querycell) |
|`Down(Optional Integer count)` | [`QueryCell`](#querycell) |
|`Left(Optional Integer count)` | [`QueryCell`](#querycell) |
|`SelectRight(Integer count, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectDown(Integer count, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectUp(Integer count, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectLeft(Integer count, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectRightTo(`[`QueryCell`](#querycell)` cell, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectDownTo(`[`QueryCell`](#querycell)` cell, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectUpTo(`[`QueryCell`](#querycell)` cell, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`SelectLeftTo(`[`QueryCell`](#querycell)` cell, Optional Boolean includeSelf)` | [`QueryRange`](#queryrange) |
|`Column()` | `Integer` |
|`Row()` | `Integer` |
|`Contains(Auto text)` | `Boolean` |
|`ContainsExactly(Auto text)` | `Boolean` |
|`ContainsNotExactly(Auto text)` | `Boolean` |
|`LessThan(Double value)` | `Boolean` |
|`GreaterThan(Double value)` | `Boolean` |
|`EqualTo(Double value)` | `Boolean` |
|`NotEmpty()` | `Boolean`  |
|`IsName(String name)` | `Boolean` |
|`Name()` | `Boolean`

##QueryColor

| Macro | Returns |
|:---------------|:--------------:|
|`RED` | `Long` |
|`BLACK` | `Long` |
|`WHITE` | `Long` |
|`LIGHT_GREEN` | `Long` |
|`GREEN` | `Long` |
|`LIGHT_BLUE` | `Long` |
|`BLUE` | `Long`

##QueryInteract

| Macro | Returns |
|:---------------|:--------------:|
|`Inform(Optional String message, Optional String title)` | [`QueryInteract`](#queryinteract) |
|`AskYesNo(Optional String message, Optional String title)` | [`QueryInteract`](#queryinteract) |
|`AskForInput(Optional String message, Optional String title, Optional String default)` | [`QueryInteract`](#queryinteract) |
|`Wait(Optional Long delay)` | [`QueryInteract`](#queryinteract) |
|`Yes()` | `Boolean` |
|`No()` | `Boolean` |
|`Valid()` | `Boolean` |
|`Value()` | `String`

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
Dim Excel: Set Excel = New Query
Dim Color: Set Color = New QueryColor

Excel.Workbook(ThisWorkbook) _
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
