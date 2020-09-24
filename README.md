<div align="center">

## Extract Numerical Values from Text Strings


</div>

### Description

The purpose of this routine is to take a string of text (such as with a textbox) and extract a numerical value from it. let's say that you have a textbox in which people enter dollar amounts. Many users are likely to enter something such as "$ 4,335.49" and expect calculations to be performed on it. The trouble is, the value of that string is 0 (zero), not 4335.49!
 
### More Info
 
The function shown below called PurgeNumericInput requires one argument. That argument is a string containing numbers with or without special characters.

Using the following function, a person would actually be able to enter a string like "$4,335.49" or even "4335.49 dollars" and still have the value returned as 4335.49.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Tips and Source Code](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-tips-and-source-code.md)
**Level**          |Unknown
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-tips-and-source-code-extract-numerical-values-from-text-strings__1-158/archive/master.zip)





### Source Code

```
Function PurgeNumericInput (StringVal As Variant) As Variant
  On Local Error Resume Next
  Dim x As Integer
  Dim WorkString As String
  If Len(Trim(StringVal)) = 0 Then Exit Function ' this is an empty string
  For x = 1 To Len(StringVal)
    Select Case Mid(StringVal, x, 1)
      Case "0" To "9", "." 'Is this character a number or decimal?
        WorkString = WorkString + Mid(StringVal, x, 1) ' Add it to the string being built
    End Select
  Next x
  PurgeNumericInput = WorkString 'Return the purged string (containing only numbers and decimals
End Function
You then just need to call the function passing a string argument to it. An example is shown below.
Sub Command1_Click
  Dim NewString as Variant
  NewString = PurgeNumericInput("$44Embedded letters and spaces 33 a few more pieces of garbage .9")
  If Val(NewString) 0 Then
    MsgBox "The Value is: " & NewString
  Else
    MsgBox "The Value is ZERO or non-numeric"
  End If
End Sub
Notice how much alphanumeric garbage was placed in the string argument. However, the returned value should be 4433.9! Two questions might arise when using this type of example.
#1 - What if the string was "0"? This could be determined by checking the length of the string (variant) returned. If the user entered a "0" then the length of the string would be > 0.
#2 - What if the string contains more than one decimal? You could use INSTR to test for the number of decimals. However, chances are, if the user entered more than one decimal you might better have them re-enter that field again anyway. <sly smile>
```

