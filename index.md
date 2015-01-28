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

## Contents

The following sections will delve deeper into the topics

### VBA Basics

<div class="posts">
  {% for post in site.categories.vba-basics reversed %}
    <a href="{{ site.home }}{{ post.url }}">{{ post.title }}</a></br>
  {% endfor %}
</div>

### ExcelQuery Basics
