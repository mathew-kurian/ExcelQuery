---
layout: page
title: ExcelQuery
tagline: Supporting tagline
---
{% include JB/setup %}

### Welcome to GitHub Pages.

This automatic page generator is the easiest way to create beautiful pages for all of your projects. Author your page content here using GitHub Flavored Markdown, select a template crafted by a designer, and publish. After your page is generated, you can check out the new branch:

### Update Author Attributes

In `_config.yml` remember to specify your own data:

{% highlight vb.net %}
Public Sub SummarizePDN()
    ' Required
    Dim Excel As Query: Set Excel = ImportExcelQuery
    Dim Worksheets As QueryBook: Set Worksheets = Excel.Workbook
    ' Set Constants
    UpdateConstants
    ' Get the search area
End Sub
{% endhighlight %}

The theme should reference these variables whenever needed.

### To-Do

This theme is still unfinished. If you'd like to be added as a contributor, [please fork](http://github.com/plusjade/jekyll-bootstrap)!
We need to clean up the themes, make theme usage guides with theme-specific markup examples.


