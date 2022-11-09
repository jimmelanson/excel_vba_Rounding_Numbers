# excel_vba_Rounding_Numbers
This method allows you to round to a specific whole number or round to a specific decimal (single).

Excel VBA has a built in Round() method that allows you to round a single/double/decimal to a specific precision (number of places after the dot) or from a decimal to the closest whole number.

What it will not let you do is round to the nearest specific whole number or round a number to the nearest specific decimal value.

For example, you can round 51.234 to 51.2 or to 51 but you cannot round it to the nearest multiple of 5.

```
Round(51.234, 1)
-> 51.2

Round(52.234, 0)
-> 51
```

The method I have created will allow you to round to something other than what Microsoft is forcing you to round to.

For example you can round currencies to the nearest EVEN amount:

```
Dim curValue As Currency
curValue = 5.47352
Debug.Print curValue
Debug.Print RoundToNumber(curValue, 0.02)
-> 5.48

curValue = 5.926
Debug.Print RoundToNumber(curValue, 0.02)
-> 5.92
```

You could round an integer to the nearest multiple of 25:
```
Dim intValue As Integer
intValue = 436
Debug.Print RoundToNumber(intValue, 25)
-> 425
```

The method can be used in VBA or it can be used as an in-cell formula:
<code>=RoundToNumber(A1, B1)</code>


