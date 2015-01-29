---
layout: page
categories : vba-basics
index: 2
title: Function Signature
---
{% include JB/setup %}

## Function Signature

Let's take the function (shown below) we wrote `Min()` function we wrote in the previous [post]({{ site.home}}/vba-basics/functions).

```vb.net
Private Function Min(firstValue As Integer, secondValue As Integer) As Integer
    If firstValue < secondValue Then
        Min = firstValue
    Else: Min = secondValue
    End If
End Function
```

The first line of a function is called a function signature (as shown below). It is used to identify the name of the function itself and the input and output data. Let's break it down and understand each component seperately.

```vb.net
Private Function Min(firstValue As Integer, 
                 secondValue As Integer) As Integer
```

### The Breakdown

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

## Striking Similarity

I am guessing you have been doing Excel before and you are here to improve your knowledge which is admirable. So you must have heard of something called `VLOOKUP`. You called `VLOOKUP` a formula, but what you might not have known is that `VLOOKUP` is actually a function but Microsoft decided to simplify the terminology from function to formula (as in "math formula"). The `VLOOKUP` function signature is as follows:

```vb.net
Public Function VLOOKUP(lookup_value As Range, table_array As Range, 
                col_index_num As Range, range_lookup As Variant)
```

As you can see there is a striking similarity to the `Min` function/formula you just wrote.

```vb.net
Private Function Min(firstValue As Integer, 
                 secondValue As Integer) As Integer
```

> ### Question
> Can you tell me why `VLOOKUP` used the `Public` keyword?
> ### Tip
> If you are confused why `VLOOKUP` uses a `Variant` for one of its parameter, refer to the [Data Types]({{ site.home }}/vba-basics/data-types) post.

