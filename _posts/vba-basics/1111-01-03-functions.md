---
layout: page
categories : vba-basics
index: 2
title: Functions
---
{% include JB/setup %}

## Functions

Functions provide a simple way to modularize algorithms. Say we have two (2) integers, `x` and `y` and we want to find the minimum number. And lets say we have to use this algorithm with three (3) sets of numbers. The subsequent sections focus on finding the miniminum between two numbers. The first section will focus on a copy-paste sort of solution; however the second section will focus on modularizing it using functions.

### Without Functions

In the following example, each comparison is done indvidiually without the use of fuctions.

```vb.net
' Entry subroutine
Public Sub GetMinOfThreeSets()

    ' Set 1
    Dim x1 As Integer : x1 = 5
    Dim y1 As Integer : y1 = 8

    ' Set 2
    Dim x2 As Integer : x2 = 9
    Dim y2 As Integer : y2 = 6

    ' Set 3
    Dim x3 As Integer : x3 = 7
    Dim y3 As Integer : x3 = 7

    ' Compare first set of numbers
    If x1 < y1 Then
        Debug.Print x1
    Else: Debug.Print y1
    End If

    ' Compare second set of numbers
    If x2 < y2 Then
        Debug.Print x2
    Else: Debug.Print y2
    End If

    ' Compare third set of numbers
    If x2 < y2 Then
        Debug.Print x2
    Else: Debug.Print y2
    End If

End Sub
```

> Run **GetMinOfThreeSets**

```
5
6
7
```

> ### Question
> What if you want to modify the logic from finding the minimum to maximum (such that you print the bigger of the two values)? Should we go and change each instance of the `<` and change it to `>`?
> You might ask that there has to better way. In fact there is! Read on to find out.

### The Functional Way

With the help of functions we can make the code shorter, more readable, and easier to modify and maintain. The following code provides the same logic as the previous example but with the use of functions.

```vb.net
' Entry subroutine
Public Sub GetMinOfThreeSets()

    ' Set 1
    Dim x1 As Integer : x1 = 5
    Dim y1 As Integer : y1 = 8

    ' Set 2
    Dim x2 As Integer : x2 = 9
    Dim y2 As Integer : y2 = 6

    ' Set 3
    Dim x3 As Integer : x3 = 7
    Dim y3 As Integer : x3 = 7

    Debug.Print Min(x1, y1)
    Debug.Print Min(x2, y2)
    Debug.Print Min(x3, y3)

End Sub

' Function which simpifies it for us
Private Function Min(firstValue As Integer, secondValue As Integer) As Integer
    If firstValue < secondValue Then
        Min = firstValue
    Else: Min = secondValue
    End If
End Function
```

> Run **GetMinOfThreeSets**

```
5
6
7
```

Let's take the function we wrote above and break it down into its parts.

```vb.net
Private Function Min(firstValue As Integer, secondValue As Integer) As Integer
    If firstValue < secondValue Then
        Min = firstValue
    Else: Min = secondValue
    End If
End Function
```

#### Function Signature

The first line of a function is called a function signature. It is used to identify the name of the function itself and the input and output data.

| **Private** | This tells the visibility of the function. The question to ask is should non-developers be able to see this when they are trying to run the macro? This specific `Min` function we are working doesn't really need to be exposed to non-developers because it is a very functional piece of code that only the developer would understand. The other option is to change it to `Public` or you can even choose to leave it empty. |
|------------:|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Function** | Required by the VBA Language standards |
|  **Min** | You need to provide a function some name. In this case, we named it `Min`. You can name it whatever you want. Just remember the name you give it is what you will be calling it when you want to use it. |
| **(** | Required by VBA Language standards |
| **firstValue** | This is considered a parameter. It is something that someone might give you. Functions are like little machines, you give it something and it will do something. In our case it is one (1) of the two (2) numbers that are required. |
| **As Integer** | The parameter we need should be a number of some sort. So, we can tell VBA that we want the input to be just that. Say for instance, we don't provide the type of input, then a user can give as any [type]({{ site.home }}/vba-basics/data-types) of data. By marking as `Integer`, we enforce the type. |
| **,** | Comma to separerate the parameters. |
| **secondValue** | The is second value that will be give to us. For our `Min` function we will do a comparison between the `firstValue` and the `secondValue`. You can give these parameters any name you want, I just named them based on their order. |
| **As Integer** | Refer to the last `As Integer` |
| **)** | Required by VBA Language standards |
| **As Integer** | This tells VBA that this function will give us back a number. |


> #### Question
> Can you make a function that multiplies three (3) numbers together?
