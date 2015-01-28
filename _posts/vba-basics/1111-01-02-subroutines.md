---
layout: page
categories : vba-basics
index: 3
title: Subroutines
---
{% include JB/setup %}

## Subroutines

Subroutines are essentially sections that help modularize lines of code. Consider the code below which shows two subroutines: `PrintHello()` and `PrintCat()`. The former only prints out `Hello World!` and the latter prints out `Cat`. Now we can choose which one we want to run.

```vb.net
' Prints only Hello World
Public Sub PrintHello()
    Debug.Print "Hello World!"
End Sub

' Prints only Cat
Public Sub PrintCat()
    Debug.Print "Cat"
End Sub
```

> Run **PrintHello**

```
Hello World!
```

> Run **PrintCat**

```
Cat
```

> ### Tip
> You can run `PrintCat()` from `PrintHello()` by doing the following.

> ```vb.net
> ' Prints both the statements
> Public Sub PrintHello()
>     Debug.Print "Hello World!"
>     Call PrintCat
> End Sub
>  ```
> Run **PrintHello**
>
> ```
> Hello World!
> Cat
> ```
