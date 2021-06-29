Attribute VB_Name = "InvoiceRound"
Public Function INVOICERND(Invoice As Range, Col As Integer, DataS As String) As Double
    Application.Volatile (False)
    
' Data Collection
    
    InvoiceS = CStr(Invoice.Value)                                                      'Get raw input
    InvArrayString = StrConv(InvoiceS, vbUnicode)                                       'Convert to unicode to allow splitting of individual characters
    InvArray = Split(Left(InvArrayString, Len(InvArrayString) - 1), vbNullChar)         'Create array of split characters
    InvLength = 0                                                                       'Set accumulator
    
    For i = LBound(InvArray) To UBound(InvArray)                                        'Loop through entire array of characters to find out when numbers end
        If IsNumeric(InvArray(i)) Then
            InvLength = InvLength + 1                                                   'InvLength then tells you how long the invoice number is
        Else
            Exit For
        End If
    Next i
            
    Length = CInt(Len(InvoiceS) - (InvLength - 6))                                      'Control flow below "thinks" invoice no. is always 6 digits, so this formula accounts for when it is not.
    CleanInv = CDbl(Left(InvoiceS, InvLength))                                          'Truncate original raw input to get just the invoice number for lookup.
    Dim Data As Variant
    Data = Range(DataS).Value                                                           'Set data range for Vlookup
    
    i = 1                                                                               'Set accumulator to look at first element in two dimensional array.
    
    Do Until (Data(i, 1) = CleanInv Or i = UBound(Data))                                'Loop until Data(i,1) matches lookup value or upper bound of data set.
        i = i + 1
    Loop
       
    PaymentRaw = Data(i, Col)                                                           'Raw payment is referenced when lookup cell is found.
    
' Payment Algorithm
    
    If (Length = 6) Then
        INVOICERND = PaymentRaw                                                         '6 digit invoice indicates no suffix, so no sum or rounding.
        Exit Function
    ElseIf (Length = 7) Then
        Oper = Right(InvoiceS, 1)                                                       '7 digit invoice indicates suffix length 1, so + or / or * is possible.
        If (Oper = "+") Then
            Payment = WorksheetFunction.Ceiling_Math(PaymentRaw, 1, 1)                  'Appropriate ceiling_math function is applied to get the desired rounding
            INVOICERND = Payment
            Exit Function
        ElseIf (Oper = "/") Then
            Payment = WorksheetFunction.Ceiling_Math(PaymentRaw, 10, 1)
            INVOICERND = Payment
            Exit Function
        ElseIf (Oper = "*") Then
            Payment = WorksheetFunction.Ceiling_Math(PaymentRaw, 0.1, 1)
            INVOICERND = Payment
            Exit Function
        End If
    ElseIf (Length = 8) Then                                                            'Length 8 indicates suffix length 2, so -# or ++ or ** possible.
        Suffix = Right(InvoiceS, 2)
        If (Right(Suffix, 1) = "+") Then
            Payment = WorksheetFunction.Ceiling_Math(PaymentRaw, 5, 1)
            INVOICERND = Payment
            Exit Function
        ElseIf (Right(Suffix, 1) = "*") Then
            Payment = WorksheetFunction.Ceiling_Math(PaymentRaw, 0.25, 1)
            INVOICERND = Payment
            Exit Function
        Else
            INVOICERND = PaymentRaw                                                     'Even if suffix is -#, no rounding is necessary, so just apply PaymentRaw.
            Exit Function
        End If
    ElseIf (Length = 9) Then                                                            'Length 9 indicates suffix of -#+, -#/, -#*, or -##
        Suffix = Right(InvoiceS, 2)
        Oper = Right(Suffix, 1)
        If ((Oper <> "+") And (Oper <> "/") And (Oper <> "*")) Then                     'If no rounding char, then just use PaymentRaw (see length 8).
            INVOICERND = PaymentRaw
            Exit Function
        ElseIf (Oper = "+") Then                                                        'Detect each rounding char and apply sum algorithm also.
            Count = CInt(Left(Suffix, 1))                                               'Read count number
            AdjCount = (Count - 1)                                                      'Adjust as to not double count current cell
            Start = Invoice.Offset(-AdjCount, 2).Address                                'Find address of first invoice in sum
            Finish = Invoice.Offset(-1, 2).Address                                      'Find address of second invoice in sum
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)     'Calculate sum then manually add current cell to avoid circular references.
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 1, 1)                          'Apply ceiling function for desired rounding character
            Adjustment = Sum - RawSum                                                   'Find difference between raw sum of all payments and rounded sum of all payments
            Payment = PaymentRaw + Adjustment                                           'Add the difference to last invoice in sum range.
            INVOICERND = Payment
            Exit Function
        ElseIf (Oper = "/") Then
            Count = CInt(Left(Suffix, 1))
            AdjCount = (Count - 1)
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 10, 1)
            Adjustment = Sum - RawSum
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        ElseIf (Oper = "*") Then
            Count = CInt(Left(Suffix, 1))
            AdjCount = (Count - 1)
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 0.1, 1)
            Adjustment = Sum - RawSum
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        End If
    ElseIf (Length = 10) Then                                                           'Length 10 indicates suffix of -#++, -##+, -##/, -#**, or -##*
        Suffix = Right(InvoiceS, 3)
        Oper = Right(Suffix, 1)
        Secondary = Left(Suffix, 2)
        MiddleChar = Right(Secondary, 1)                                                'In addition to normal formatting to isolate suffix, must find middlecharacter to distinguish between -#++ and -##+ (or -#** and -##*)
        If (Oper = "/") Then
            Count = CInt(Secondary)
            AdjCount = Count - 1
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 10, 1)
            Adjustment = Sum - RawSum
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        ElseIf (MiddleChar = "+") Then
            Count = CInt(Left(Suffix, 1))
            AdjCount = Count - 1
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 5, 1)
            Adjustment = Sum - RawSum
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        ElseIf (MiddleChar = "*") Then
            Count = CInt(Left(Suffix, 1))
            AdjCount = Count - 1
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 0.25, 1)
            Adjustment = Sum - RawSum
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        Else                                                                            'If middlechar is a number then only cases -##+ and -##* exist.
            Count = CInt(Secondary)
            AdjCount = Count - 1
            Start = Invoice.Offset(-AdjCount, 2).Address
            Finish = Invoice.Offset(-1, 2).Address
            RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
            Sum = WorksheetFunction.Ceiling_Math(RawSum, 1, 1)                          'Calculate sum for -##+
            DecSum = WorksheetFunction.Ceiling_Math(RawSum, 0.1, 1)                     'Calculate sum for -##*
            If (Oper = "+") Then                                                        'Apply control flow to determine which sum to use.
                Adjustment = Sum - RawSum
            Else
                Adjustment = DecSum - RawSum
            End If
            Payment = PaymentRaw + Adjustment
            INVOICERND = Payment
            Exit Function
        End If
    ElseIf (Length = 11) Then                                                           'Length 11 indicates suffix of -##++ or -##**
        Suffix = Right(InvoiceS, 4)
        Oper = Right(Suffix, 1)
        Count = CInt(Left(Suffix, 2))
        AdjCount = Count - 1
        Start = Invoice.Offset(-AdjCount, 2).Address
        Finish = Invoice.Offset(-1, 2).Address
        RawSum = CDbl(WorksheetFunction.Sum(Range(Start, Finish)) + PaymentRaw)
        Sum = WorksheetFunction.Ceiling_Math(RawSum, 5, 1)
        DecSum = WorksheetFunction.Ceiling_Math(RawSum, 0.25, 1)
        If (Oper = "+") Then
            Adjustment = Sum - RawSum
        Else
            Adjustment = DecSum - RawSum
        End If
        Payment = PaymentRaw + Adjustment
        INVOICERND = Payment
        Exit Function
    End If
    
End Function







