---
layout: page
title: ExcelQuery
tagline: Supporting tagline
---
{% include JB/setup %}

## Welcome to the ExcelQuery Tutorials

The tutorial will start with some basic VBA concepts and then introduce ExcelQuery. After the essential queries are covered more complex algorithms will be discussed. The tutorial will be detailed in order to ensure all questions are answered. If you still have any questions, feel free to shoot me an [email](mailto:bluejamesbond@gmail.com).

## What is VBA?

VBA is a programming language used to access the document models of an Excel workbook. VBA applications running on Excel will have access to the sheets, range, and cells of your file(s). Applications can also access system files such as your documents, files, music, etc. So in essence, VBA is capable of doing a wide variety of things.

> ### Did you know?
> Microsoft Excel and VBA itself are actually built using a much more powerful language called C++. The C++ programming language provides developers with tools that enable them to reduce the memory consumption and increase performance. However, such performance is rarely required when programming directly with Excel sheets, ranges, and cells.

### VBA Basics

<div class="posts">
  {% for post in site.categories.vba-basics reversed %}
    <a href="{{ site.home }}{{ post.url }}">{{ post.title }}</a></br>
  {% endfor %}
</div>

## What is ExcelQuery

ExcelQuery is a programming library (and framework) which is built on top of VBA. ExcelQuery contains frequent algorithms that VBA developers use. It also implements a way to spot and identify your coding errors. This feature makes coding easier for beginners. ExcelQuery wants developers to *write less and do more.*

> ### Fun Fact
> ExcelQuery is actually based off JQuery, a library for working with web applications (like websites). In the examples below, we are performing similar tasks but one is with JQuery while the other with ExcelQuery. Can you follow the both pieces of code?
> #### JQuery
> ```js
$(".box:not(:empty)")              // In websites, we have boxes
       .eq(1)                      // Select the first one
       .css("foreground","red")    // Set text color to red
       .css("background", "blue")  // Set background to red
       .value(100)                 // Update value to 100
```
> #### ExcelQuery
> ```vb.net
Excel.Workbook.Worksheet("MyWorksheet")
       .Cells("A1:A40") _              ' In excel, we have cells
       .SelectNotEmpty _
       .Eq(1) _                        ' Select the first one
       .Foreground(Excel.Color.RED) _  ' Set text color to red
       .Foreground(Excel.Color.BLUE) _ ' Set background to red
       .Update(100)                    ' Update value to 100
```
> You can learn more about JQuery [here](http://jquery.com/).

### ExcelQuery Basics

Coming soon


