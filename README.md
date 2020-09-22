<div align="center">

## \(Updated Again\) Check for duplicates in an array\!


</div>

### Description

I've been looking around for this code and no one could provide it. So finally I wrote it. It checks for duplicates in an array and returns true if there are any.
 
### More Info
 
The only input required is the array

You could use this code to generate lottery numbers or check if more than one record of the same name is present...

Returns True if there are any duplicates and false otherwise

If used with LARGE (and I mean LARGE as in arrays with hundreds or thousands of items) it will slow down. USE WITH ONE DIMENSIONAL ARRAYS


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Arthur Chaparyan3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arthur-chaparyan3.md)
**Level**          |Unknown
**User Rating**    |3.5 (28 globes from 8 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/arthur-chaparyan3-updated-again-check-for-duplicates-in-an-array__1-2856/archive/master.zip)





### Source Code

```
Public Function AnyDup(NumList As Variant) As Boolean
 Dim a As Long, b As Long
 'Start the first loop
 For a = LBound(NumList) To UBound(NumList)
 'Start the second loop (thanks for the suggestions everyone)
 For b = a + 1 To UBound(NumList)
 'Check if the values are the same
 'if they're equal, then we found a duplicate
 'tell the user and end the function
 If NumList(a) = NumList(b) Then AnyDup = True: Exit Function
 Next
 Next
End Function
```

