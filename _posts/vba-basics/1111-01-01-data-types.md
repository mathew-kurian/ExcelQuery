---
layout: page
categories : vba-basics
index: 3
title: Data Types
---
{% include JB/setup %}

## General VBA Data Types

VBA developers can choose to declare the type of data that is being being used whether in a [`Dim`]({{ site.home }}/vba-basics/dim) or paramters to a [`Function`]({{ site.home }}/vba-basics/functions). Variables are declared to be a certain type using the `As` keyword. It ensures that the data will be of a certain type.

### Primary Primitive Data Types

|  **String** | String represents a _string of characters_ or text. Strings are usually initialized with quotes i.e. `"this is a string of data"`. It is also important note that an empty String is defined as `""`. |
|------------:|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|  **Double** | Double represents a _decimal number_. They can also be __whole numbers_. Numbers such as `-1.11`, `0`, `9.4`, `5`, `10`, and even `10000000` are considered as doubles.                                                                                |
| **Boolean** | Booleans are represent truth values namely `true` or `false`. Booleans can only be in one state at a time and they only have two total states.                                                        |

### Secondary Primitive Data Types

| **Integer** | Integers represents _whole numbers_ between `-32,788` and `32,787`.  **Question:** What is the difference between an `Integer` and `Double`?                                                                        |
|------------:|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|    **Long** | Long represents _whole numbers_ between `-2,147,483,648` and `2,147,483,647`. Longs support much larger numbers than the aforementioned Integers so they use up more memory on the computer and tend to be slower. |
| **Variant** | Variants are tricky. To put it simply, Variants mean `Anything`. In other words, something marked `As Variant` could be anything! This is why they can be confusing. |

### Complex Data Types

These data types are composed of a group of the primitive data types explained above. One of the most important complex data type is defined below. There are several more complex data types which I have not mentioned. i will introduce them to you later in the tutorial.

| **Collections** | Collections represent a list of data-types. An example of a collection is `[ 1, "data", 1.2, "34" ]`. The important thing is that order does matter. So the element at index 1 in this case is `1` and the element at index 2 is `"data"` an so on. You can usually get the length of the Collection which at times is proves to be very useful.  |
|--------------:|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
