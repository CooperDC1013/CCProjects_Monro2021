Attribute VB_Name = "InvoiceSum"
Private Start As Range
Private Finish As Range
Private RawSum As Double
Private Invoice As Range
Private AdjCount As Integer
Private SumCol As Integer

Private Sub Assign()                                                                    'Sub sets the beginning and ending cells of the range of invoices to be summed.
        
        Set Start = Invoice.Offset(-AdjCount, SumCol)
        Set Finish = Invoice.Offset(0, SumCol)
        RawSum = CDbl(WorksheetFunction.Sum(Range(Start.Address, Finish.Address)))
    
End Sub

Public Function INVOICESM(Inp As Range, Optional myAutoColl As Integer = 1, Optional SumColOff As Integer = 4) As Variant
    Application.Volatile (False)
    
    SumCol = SumColOff
    AutoCol = myAutoColl
    
    Set Invoice = Inp                                                                   'Take invoice from inp as to preserve parameter variable
    InvoiceS = CStr(Invoice.Value)                                                      'As String from range
    
    SumCol = SumColOff - Invoice.Column
    
    InvArrayString = StrConv(InvoiceS, vbUnicode)                                       'Create string in Unicode for character splitting
    InvArray = Split(Left(InvArrayString, Len(InvArrayString) - 1), vbNullChar)         'Split into individual characters into an array to deduce actual invoice number
    InvLength = 0                                                                       'Assign accumulator variable for invoice true length
    
    For i = LBound(InvArray) To UBound(InvArray)                                        'Loop until next character is not numeric, thus determining length of invoice number
        If IsNumeric(InvArray(i)) Then
            InvLength = InvLength + 1
        Else
            Exit For
        End If
    Next i
    
    Length = CInt(Len(InvoiceS) - (InvLength - 6))                                      'Control flow below "thinks" invoice number is always 6 digits with a suffix attached, this formula accounts for when it is not.
    Suffix = WorksheetFunction.Replace(InvoiceS, 1, InvLength, "")                      'Replace the invoice digits with "" to isolate suffix appended.
    Key = Left(Suffix, 1)                                                               'Key is "-" if a suffix exists.
    Dim Run As Boolean
    
    If Key = "-" Then                                                                   'Only run sum algorithm if key is present.
        Run = True
        Code = WorksheetFunction.Replace(Suffix, 1, 1, "")                              'Remove "-"
    
    ElseIf Key = "." Then
        
        Account = Cells(Invoice.Row, AutoCol).Value
        'Account = Invoice.Offset(0, 1 - Invoice.Column).Value
        TargetRow = Invoice.Row
        Target = Cells(TargetRow, AutoCol).Value
        'Target = Invoice.Offset(TargetRow, 1 - Invoice.Column)
        
        Count = 0
        
        Do While Target = Account
        
            Count = Count + 1
            TargetRow = TargetRow - 1
            Target = Cells(TargetRow, AutoCol).Value
            'Target = Invoice.Offset(TargetRow, 1 - Invoice.Column)
            
            If IsError(Target) Then
                Target = 0
            End If
            
        Loop
        
        AdjCount = Count - 1
        Assign
        INVOICESM = RawSum
        Exit Function
    
    End If
    
    If Run = True Then
    
        If Length = 8 Then                                                              'When key is present, Length 8 is -# so code is #
            
            Count = CInt(Right(Code, 1))
            AdjCount = Count - 1
            Assign
            INVOICESM = RawSum
            Exit Function
        
        ElseIf Length = 9 Then                                                          'Length 9 will always have a sum but if next if passes then it's 2 digit sum
        
            If (Right(Code, 1) <> ("+")) And (Right(Code, 1) <> ("/")) And (Right(Code, 1) <> ("*")) Then
                Count = CInt(Right(Code, 2))
                AdjCount = Count - 1
                Assign
                INVOICESM = RawSum
                Exit Function
            
            ElseIf Right(Code, 1) = "+" Then                                            'All ElseIfs are for 1 digit sum with rounding operator.
            
                Count = CInt(Left(Code, 1))
                AdjCount = Count - 1
                Assign
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 1, 1))
                Exit Function
            
            ElseIf Right(Code, 1) = "/" Then
                
                Count = CInt(Left(Code, 1))
                AdjCount = Count - 1
                Assign
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 10, 1))
                Exit Function
            
            ElseIf Right(Code, 1) = "*" Then
                
                Count = CInt(Left(Code, 1))
                AdjCount = Count - 1
                Assign
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 0.1, 1))
                Exit Function
                
            End If
        
        ElseIf Length = 10 Then                                                         'Observe middlechar to determine one or two digit sum.
        
            MiddleChar = Left(WorksheetFunction.Replace(Code, 1, 1, ""), 1)
            
            If MiddleChar = "+" Then
                
                Count = CInt(Left(Code, 1))
                AdjCount = Count - 1
                Assign
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 5, 1))
                Exit Function
                
            ElseIf MiddleChar = "*" Then
            
                Count = CInt(Left(Code, 1))
                AdjCount = Count - 1
                Assign
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 0.25, 1))
                Exit Function
            
            Else                                                                       'Operations for two digit sum
            
                Oper = Right(Code, 1)
                Count = CInt(Left(Code, 2))
                AdjCount = Count - 1
                
                If Oper = "+" Then
                    
                    Assign
                    INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 1, 1))
                    Exit Function
                
                ElseIf Oper = "/" Then
                    
                    Assign
                    INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 10, 1))
                    Exit Function
                    
                ElseIf Oper = "*" Then
                
                    Assign
                    INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 0.1, 1))
                    Exit Function
                
                End If
            
            End If
        
        ElseIf Length = 11 Then                                                             'Two digit sum with only ++ or ** rounding operators
        
            Count = CInt(Left(Code, 2))
            AdjCount = Count - 1
            Assign
            
            If Right(Code, 1) = "+" Then
            
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 5, 1))
                Exit Function
            
            Else
            
                INVOICESM = CDbl(WorksheetFunction.Ceiling_Math(RawSum, 0.25, 1))
                Exit Function
            
            End If
            
        End If
    
    Else
        
        INVOICESM = ""
        Exit Function
    
    End If
        
End Function

