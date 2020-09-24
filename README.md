<div align="center">

## formatcurrency


</div>

### Description

format a text field into a $ currency field.
 
### More Info
 
text box name

Call teh module by passing any text box name that you want to have as a currency text box.

formatted value in $


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vikramjit Singh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vikramjit-singh.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vikramjit-singh-formatcurrency__1-1441/archive/master.zip)

### API Declarations

```
'put as many text boxes as you want...say text1
'call the class by using the following line
Call validate(Text1)
```


### Source Code

```
Sub validate(textboxname As TextBox)
If KeyAscii > Asc(9) Or KeyAscii < Asc(0) Then
KeyAscii = 0
End If
mypos = InStr(1, textboxname.Text, ".")
If mypos <> 0 Then
textboxname.Text = Format(textboxname.Text, "$###,###,###,###.##")
Else
textboxname.Text = Format(textboxname.Text, "$###,###,###,###")
End If
End Sub
```

