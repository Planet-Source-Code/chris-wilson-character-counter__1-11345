<div align="center">

## Character counter


</div>

### Description

A FAST and EASY way to count the number of occurrences of one string within another! Please vote for this tip if you find it useful.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-wilson.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-wilson-character-counter__1-11345/archive/master.zip)





### Source Code

```
Private Sub Test()
Const mystr = "This is a test of the split function"
' returns 6
Debug.Print Occurs(mystr, "t")
End Sub
Public Function Occurs(ByVal strtochk As String, ByVal searchstr As String) As Long
' remember SPLIT returns a zero-based array
Occurs = UBound(Split(strtochk, searchstr)) + 1
End Function
```

