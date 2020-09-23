<div align="center">

## A Binary Printing with only 10 lines of code \!


</div>

### Description

This procedure prints the binary code of a number on Immediate window.
 
### More Info
 
Num as Long

You can easily change the procedure to a function for returning the binary code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eduardo B\. Castro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eduardo-b-castro.md)
**Level**          |Advanced
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eduardo-b-castro-a-binary-printing-with-only-10-lines-of-code__1-55545/archive/master.zip)





### Source Code

```
Public Sub PrintBinary(Num As Long)
 Dim j&, i&
 j = 128
 For i = 8 To 1 Step -1
 If (Num And j) = 0 Then
  Debug.Print "0";
 Else
  Debug.Print "1";
 End If
 j = j / 2
 Next
End Sub
```

