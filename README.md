<div align="center">

## GetDaysInMonth


</div>

### Description

'This function will use the computers clock to retrive the days in any month, including leap years.

'If you use this code then please leave your comments or even vote if you have time..
 
### More Info
 
'The are only two inputs one of which is optional.

'The First Is ValMonth which is the month you want to get the days for.

'The second is the year, if not obmited then the current year is used.

'This code is in a function, if you are not sure about functions then please read about them before using this code.

'The days in the month specifyed

'None that i know about, and there souldnt be any


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cux](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cux.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cux-getdaysinmonth__1-40285/archive/master.zip)





### Source Code

```
'Please dont remove this TAG
'By CoderUX 31-10-2002
'Copy this code into your form or module
Public Function GetDaysInMonth(ByVal ValMonth As Double, Optional ByVal ValYear As Double) as double
On Error GoTo Handler:
 If ValYear = Empty Then
 GetDaysInMonth = DateDiff("D", "01/" & ValMonth & "/" & Format(Now, "YYYY"), DateAdd("M", 1, "01/" & ValMonth & "/" & Format(Now, "YYYY")))
 Else
 GetDaysInMonth = DateDiff("D", "01/" & ValMonth & "/" & ValYear, DateAdd("M", 1, "01/" & ValMonth & "/" & ValYear))
 End If
Handler:
Exit Function
End Function
```

