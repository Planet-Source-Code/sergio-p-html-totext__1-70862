<div align="center">

## HTML toText


</div>

### Description

This little piece of code will strip all HTML tags from a web page. What your left with is pure text.

I've seen a few approaches, one which took up 500 lines of code. And some guy is selling an HTML stripper for $250. Rediculous.

This is my first post here. I just wanted to share and save someone some time.
 
### More Info
 
Just paste it into your form or module and call it. Text1 = HTML2Text(HTMLString1)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sergio P](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sergio-p.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sergio-p-html-totext__1-70862/archive/master.zip)





### Source Code

```
Public Function HTML2Text(ByVal HTML As String) As String
Dim X As Long
Dim B As String
Dim String1 As String
Dim Counter As Long
X = 1
B$ = Mid(HTML, X, 1)
While Len(B$) = 1
B$ = Mid(HTML, X, 1)
If B$ = "<" Then Counter = Counter + 1
If Counter = 0 Then String1$ = String1$ + B$
If B$ = ">" And Counter <> 0 Then Counter = Counter - 1
X = X + 1
Wend
HTML2Text = String1
End Function
```

