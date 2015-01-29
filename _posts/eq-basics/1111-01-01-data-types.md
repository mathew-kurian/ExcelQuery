---
layout: page
categories : eq-basics
index: 3
title: Data Types
---
{% include JB/setup %}

## ExcelQuery Data Types

Similar to the data types discussed [here]({{ site.home }}/vba-basics/data-types), ExcelQuery also has its own set of types. It is of utmost importance that you know about these data types in order for you to fully utilize the library's potential.

|  **QueryBook** | QueryBook represents a single Excel workbook. |
|------------:|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  **QuerySheet** | QuerySheet represents a single Excel sheet. |
| **QueryRange** | QueryRange represents a group of Excel cells. The cells don't have to be contiguous. It can also represent an empty set of cells (of length `0`). These cells can come from any book and any worksheet. |
| **QueryCell** | QueryCell represents exactly one (1) Excel cell. This cell can come from any book and any worksheet. |

### Important Rules
At this point, these data types may not mean much to you but you need to realize heirarchy and relationships between them.

| **#1** | One (1) `QueryBook` can contain inside it anywhere between `0` and `X` number of `QuerySheets`. |
|-------:|-------------------------------------------------------------------------------------------------|
| **#2** | One (1) `QuerySheet` can contain inside it anywhere between `0` and `Y` number of `QueryRange`. |
| **#3** | One (1) `QueryRange` can contain inside it anywhere between `0` and `Z` number of `QueryCells`. |
| **#4** | There will never be a negative number of elements.                                              |

### In Code
The goal of the programming with this library is to extract a particular a element from its parent element and then perform some task with it. I can show you a sneak peak of what I mean in the below code. The order of the data types is **critical** and it absolutely essential you understand.

```vb.net
Q.Workbook(ThisWorkbook) _  ' Get the QueryBook.
 .Worksheet("Sheet1") _     ' Get the QuerySheet. Proof of Rule #1.
 .Cells("A1:A2") _          ' Get the QueryRange. Proof of Rule #2.
 .Eq(1) _                   ' Get the QueryCell. Proof of Rule #3.
 .Update("Hello!")          ' Update the cell
```

