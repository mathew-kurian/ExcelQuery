---
layout: page
categories : basics
index: 2
title: Functions
---
{% include JB/setup %}

## Functions

Functions are the workers of the

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
