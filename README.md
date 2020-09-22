<div align="center">

## Trim all those EXTRA spaces


</div>

### Description

VB forgot to add a Function strip out all those LEADING, TRAILING and EXTRA spaces in ONE Function.

I have seen many attempts at doing this but think mine does it in the least amount of code.

Note:<spc> = literal SPACE
 
### More Info
 
String eg:Strip<spc><spc>all<spc>

Strip<spc>all


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Gillham](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-gillham.md)
**Level**          |Beginner
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-gillham-trim-all-those-extra-spaces__1-25713/archive/master.zip)





### Source Code

```
Public Function TrimALL(ByVal TextIN As String) As String
 TrimALL = Trim(TextIN)
 While InStr(TrimALL, String(2, " ")) > 0
 TrimALL = Replace(TrimALL, String(2, " "), " ")
 Wend
End Function
```

