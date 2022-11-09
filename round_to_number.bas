Attribute VB_Name = "round_to_number"
Option Explicit

'-------------------------------------------------------------
'
'       IMPORTANT NOTE
'
'   This module has two methods that are both needed. They
'   are RoundToNumber() and getPrecision().
'
'   Make sure you copy both if if you are copy and pasting
'   this code into your workbook.
'
'-------------------------------------------------------------

Private Sub RoundToNumber_Test()
    Dim i As Long
    
    Dim sngIncrementor As Single
    sngIncrementor = 0.01
    
    Dim sngStart As Single
    sngStart = 4.8
    
    Dim sngRoundTo As Single
    sngRoundTo = 0.004

    For i = 1 To 75
        Debug.Print "base " & Round(sngStart, 2) & ": " & RoundToNumber(Round(sngStart, 2), sngRoundTo)
        sngStart = sngStart + sngIncrementor
    Next i
    Debug.Print
End Sub

Public Function RoundToNumber(ByVal sngInputValue As Single, ByVal sngRoundTo As Single) As Double
    'Default value for return in case we crap out of this procedure.
    RoundToNumber = sngInputValue
    'Get the smallest precision of the two numbers (this will be the larger precision number).
    Dim longPrecision As Integer
    longPrecision = getPrecision(sngRoundTo)
    If getPrecision(sngInputValue) > longPrecision Then
        longPrecision = getPrecision(sngInputValue)
    End If
    'Make sure that both the inputvalue and the tolerance have a value greater than zero.
    If sngRoundTo > 0 Then
        'If number is negative, keep track of that to make it negative at the end.
        Dim boolIsNegative As Boolean
        boolIsNegative = False
        If sngInputValue < 0 Then
            boolIsNegative = True
            sngInputValue = Abs(sngInputValue)
        End If
        'Declare variables
        Dim longInputValue As Long
        Dim longRoundTo As Long
        Dim longMultiplier As Long
        Dim i As Long
        
        'We need to convert the inputs to longs so that we can SAFELY process
        'them with VBA math, as math with singles and doubles in VBA is notoriously
        'difficult, especially when it comes to rounding floating-point numbers.
        'https://learn.microsoft.com/en-us/office/troubleshoot/excel/floating-point-arithmetic-inaccurate-result
        'https://bettersolutions.com/vba/numbers/floating-point-numbers.htm
        'https://www.vitoshacademy.com/excel-vba-floating-point-numbers-in-excel-not-exact/
        'https://newtonexcelbach.com/2021/04/05/floating-point-precision-problems/
        
        If getPrecision(sngInputValue) = 0 And getPrecision(sngRoundTo) = 0 Then
            'Neither the input value or the round-to value has a decimal, so just
            'coerce them into longs.
            longInputValue = CLng(sngInputValue)
            longRoundTo = CLng(sngRoundTo)
            RoundToNumber = longInputValue
        ElseIf getPrecision(sngRoundTo) > 0 Then
            'The round-to is a decimal and the input value may or may not be.
            longMultiplier = 1
            For i = 1 To longPrecision
                longMultiplier = longMultiplier * 10
            Next i
            longInputValue = CLng(Round(sngInputValue * longMultiplier, longPrecision))
            longRoundTo = CLng(Round(sngRoundTo * longMultiplier, longPrecision))
        Else
            'The input value is a decimal and the round-to is not. In this case, the
            'input value needs to be rounded to a whole number as the rounded value
            'will never be a decimal if the round-to value is not a decimal.
            longInputValue = CLng(Round(sngInputValue, 0))
            longRoundTo = CLng(sngRoundTo)
        End If

        'We can exit quickly if the input value is already a multiple of the round-to value.
        'We can relatively safely process this as a Mod() because we already converted the
        'decimal values to whole numbers.
        If longInputValue Mod longRoundTo = 0 Then
            RoundToNumber = longInputValue
            If longMultiplier > 0 Then
                RoundToNumber = RoundToNumber / longMultiplier
            End If
            If boolIsNegative = True Then
                RoundToNumber = 0 - RoundToNumber
            End If
            Exit Function
        End If
        'Declare more variables and iterators if we have not exited.
        Dim longForLimit As Long
        Dim i_Before As Integer
        i_Before = 0
        Dim i_After As Integer
        i_After = 0
        'Find the closest mod 0 below the input number
        For i = longInputValue To 0 Step -1
            If i_Before = 0 Then
                If i Mod longRoundTo = 0 Or i = 0 Then
                    i_Before = longInputValue - i
                End If
            End If
        Next i
        'Find the closest mod 0 above the input number
        If longInputValue > longRoundTo Then
            longForLimit = longInputValue * 2
        Else
            longForLimit = longRoundTo * 2
        End If
        For i = longInputValue To longForLimit
            If i_After = 0 Then
                If i Mod longRoundTo = 0 Then
                    i_After = i - longInputValue
                End If
            End If
        Next i
        'Now we look to see which rounding result is closer and that will be
        'the value that is returned by the method. Here is an example, eg:
        '   Input Value: 33
        '   Rount To:    12
        '   The user wants the closest multiple of 12 to the input value of 33.
        '   Lower: We keep subtracting 1 from 33 until we find a multiple of 12:
        '       33 -> 32 -> 31 -> 30 -> 29 -> 28 -> 27 -> 26 -> 25 -> 24 = 9 Steps
        '   Higher: We keep adding 1 to 33 until we find a multiple of 12:
        '       33 -> 34 -> 35 -> 36 = 3 Steps
        '   We now see that the steps above the input value (3) is smaller than the
        '   steps below the input value (9), therefore we are going to move towards
        '   the smaller difference (above/higher) - which is also to say that we are
        '   rounding up.
        'NOTE: This explanation is consistent for whole numbers and decimal values.
        If i_After < i_Before Then
            'Rounding UP as the steps higher than the input value are smaller than
            'the steps below the input vale.
            RoundToNumber = longInputValue + i_After
        Else
            'Rounding DOWN. Vice versa of what I said above.
            RoundToNumber = longInputValue - i_Before
        End If
        If longMultiplier > 0 Then
            RoundToNumber = Round((RoundToNumber / longMultiplier), longPrecision)
        End If
        If boolIsNegative = True Then
            RoundToNumber = 0 - RoundToNumber
        End If
    End If
 End Function

Public Function getPrecision(ByVal strInputValue As String) As Integer
    If InStr(strInputValue, ".") > 0 Then
        If Len(strInputValue) - InStr(strInputValue, ".") > 0 Then
            Dim longNumber As Long
            longNumber = CLng(Right(strInputValue, Len(strInputValue) - InStr(strInputValue, ".")))
            If longNumber <> 0 Then
                getPrecision = Len(strInputValue) - InStr(strInputValue, ".")
            End If
        Else
            getPrecision = 0
        End If
    End If
End Function

