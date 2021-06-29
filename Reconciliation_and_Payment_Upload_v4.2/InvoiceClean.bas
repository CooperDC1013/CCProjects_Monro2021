Attribute VB_Name = "InvoiceClean"
Public Function INVOICECLN(Invoice As Range) As Double
    Application.Volatile (False)
        
    InvoiceS = CStr(Invoice.Value)                          'Convert to String
    InvArrayStr = StrConv(InvoiceS, vbUnicode)              'Convert to Unicode String
    InvArray = Split(Left(InvArrayStr, Len(InvArrayStr) - 1), vbNullChar) 'Convert to array of characters
    InvLength = 0                                           'Set accumulator length as 0
    
    For i = LBound(InvArray) To UBound(InvArray)            'Loop through array until no more numeric characters
        If IsNumeric(InvArray(i)) Then
            InvLength = InvLength + 1                       ' Length +1 when next char is numeric
        Else
            Exit For
        End If
    Next i
    
    INVOICECLN = CDbl(Left(Invoice.Value, InvLength))

End Function

