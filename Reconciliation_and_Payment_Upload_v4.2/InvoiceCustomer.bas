Attribute VB_Name = "InvoiceCustomer"
Public Function INVOICECUST(Inp As Range, Col As Integer, DataS As String) As Double
    Application.Volatile (False)
    
    InvoiceS = CStr(Inp.Value)                                      'Get string of lookup value
    
    InvArrayString = StrConv(InvoiceS, vbUnicode)                   'Convert string to Unicode
    InvArray = Split(Left(InvArrayString, Len(InvArrayString) - 1), vbNullChar) 'Split into characters
    InvLength = 0                                                   'Set accumulator
    
    For i = LBound(InvArray) To UBound(InvArray)                    'Loop through array of characters to determine how long true invoice number is.
        If IsNumeric(InvArray(i)) Then
            InvLength = InvLength + 1
        Else
            Exit For
        End If
    Next i
    
    CleanInv = CDbl(Left(InvoiceS, InvLength))                      'True invoice number is isolated by using InvLength
    Dim Data As Variant
    Data = Range(DataS).Value                                       'Data table is converted from string to range
    
    i = 1                                                           'Set accumulator for array index
    
    Do Until (Data(i, 1) = CleanInv Or i = UBound(Data))            'Loop through data array until lookup value of invoice number is found or end of data is reached
        i = i + 1
    Loop
    
    Dim CustRaw As String
    Dim CustNo As String
    CustRaw = CStr(Data(i, Col))                                    'Pull customer number out
    
    If IsNumeric(Right(CustRaw, 1)) Then                            'Ensure that at least one digit at end is numeric to avoid infinite looping
    
        Do While (Not (IsNumeric(Left(CustRaw, 1))))                    'Loop through characters. Keep truncating cust number as many times from front until first char is number.
            CustRaw = WorksheetFunction.Replace(CustRaw, 1, 1, "")
        Loop
    
        CustNo = CustRaw
    
    Else
    
        CustNo = ""
        
    End If
    
    Do While (Left(CustNo, 1) = "0")                                'Loop through again to erase zeros in front.
        CustNo = WorksheetFunction.Replace(CustNo, 1, 1, "")
    Loop
    
    INVOICECUST = CDbl(CustNo)
    
End Function
