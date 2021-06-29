VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AgingForm 
   Caption         =   "Aging Analysis Task Selection"
   ClientHeight    =   11280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "AgingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AgingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tilde As Integer
Public Confirm As Integer

Private Sub Cancel_Click()
    
    Dim RunText As String
    Dim V() As Variant
    
    If Tilde = 7 Then
        RunText = "Nothin' to see here pal."
    Else
        RunText = "CONFIRM RUN"
    End If
    
    Tilde = 0
    
    Dim Q As Integer
    Dim QQ As Integer
    
    If CreditHoldCheckbox.Value = True Then
        Q = 11
        QQ = 1
    Else
        Q = 9
        QQ = 0
    End If
    
    V = Array(AcctBox, DocTypeBox, InvoiceBox, InvDateBox, DueDateBox, OpenAmtBox, GrossAmtBox, BUBox, BU3Box, BU5Box, CustAcctBox, TempCreditBox)
    t = Array(DetailTabBox, CustTabBox)
    
    If Confirm = 0 Then
    
        c = ColumnConverter(V)
        
        Valid = True
        
        For i = 0 To Q
            If c(i) = "!" Then
                Valid = False
            End If
        Next i
        
        For i = 0 To QQ
        
            If IsNumeric(t(i).Value) And (t(i).Value > 0) And (t(i).Value <= Application.Worksheets.Count) Then
                If Sheets(CInt(t(i).Value)).Visible = False Then
                    Valid = False
                End If
            Else
                Valid = False
            End If
            
        Next i
        
        If Valid Then
        
            DetailTabBox.BackColor = &H80FF80
            CustTabBox.BackColor = &H80FF80
            RunRFCheckbox.Locked = True
            MWCheckbox.Locked = True
            SpreadCheckBox.Locked = True
            CreditHoldCheckbox.Locked = True
            DetailTabBox.Locked = True
            CustTabBox.Locked = True
            RunRFCheckbox.TabStop = False
            MWCheckbox.TabStop = False
            SpreadCheckBox.TabStop = False
            CreditHoldCheckbox.TabStop = False
            DetailTabBox.TabStop = False
            CustTabBox.TabStop = False
            
            For i = 0 To UBound(c)
            
                If c(i) = "?" Then
                    V(i).BackColor = &H1FFFF
                Else
                    V(i).BackColor = &H80FF80
                End If
                
                V(i).Locked = True
                V(i).TabStop = False
            
            Next i
            
'            AcctBox.BackColor = &H80FF80
'            DocTypeBox.BackColor = &H80FF80
'            InvoiceBox.BackColor = &H80FF80
'            InvDateBox.BackColor = &H80FF80
'            DueDateBox.BackColor = &H80FF80
'            OpenAmtBox.BackColor = &H80FF80
'            CustAcctBox.BackColor = &H80FF80
'            TempCreditBox.BackColor = &H80FF80
'            AcctBox.Locked = True
'            DocTypeBox.Locked = True
'            InvoiceBox.Locked = True
'            InvDateBox.Locked = True
'            DueDateBox.Locked = True
'            OpenAmtBox.Locked = True
'            CustAcctBox.Locked = True
'            TempCreditBox.Locked = True
            
            Confirm = 1
            AgingForm.Height = AgingForm.Height + 36
            Run.Visible = True
            Run.Caption = RunText
            Cancel.Caption = "RETURN"
            
            Run.TabStop = True
            Cancel.TabIndex = 0
            Run.TabIndex = 1
            
'            AcctBox.TabStop = False
'            DocTypeBox.TabStop = False
'            InvoiceBox.TabStop = False
'            InvDateBox.TabStop = False
'            DueDateBox.TabStop = False
'            OpenAmtBox.TabStop = False
'            CustAcctBox.TabStop = False
'            TempCreditBox.TabStop = False
                
        Else
            
            For i = 0 To UBound(V)
                If c(i) = "!" And (i <= Q) Then
                    V(i).BackColor = &HFF
                ElseIf c(i) = "?" Then
                    V(i).BackColor = &H1FFFF
                Else
                    V(i).BackColor = &H80FF80
                End If
                
                V(i).Locked = True
                V(i).TabStop = False
            Next i
            
            RunRFCheckbox.Locked = True
            MWCheckbox.Locked = True
            SpreadCheckBox.Locked = True
            CreditHoldCheckbox.Locked = True
            
            RunRFCheckbox.TabStop = False
            MWCheckbox.TabStop = False
            SpreadCheckBox.TabStop = False
            CreditHoldCheckbox.TabStop = False
            DetailTabBox.TabStop = False
            CustTabBox.TabStop = False
            Cancel.TabIndex = 0
            
            Confirm = 2
            Cancel.Caption = "RETURN"
            
            For i = 0 To 1
        
                If IsNumeric(t(i).Value) And (t(i).Value > 0) And (t(i).Value <= Application.Worksheets.Count) Then
                
                    t(i).BackColor = &H80FF80
                    
                    If Sheets(CInt(t(i).Value)).Visible = False Then
                        t(i).BackColor = &HFF
                    End If
                Else
                    t(i).BackColor = &HFF
                End If
                
                t(i).Locked = True
                
                If i > QQ Then
                    t(i).BackColor = &H80FF80
                End If
                
            Next i
            
        End If
        
    ElseIf Confirm = 1 Then
    
        DetailTabBox.BackColor = &H80000005
        CustTabBox.BackColor = &H80000005
        
        For i = 0 To UBound(V)
        
            V(i).BackColor = &H80000005
            V(i).Locked = False
            V(i).TabStop = True
        
        Next i
        
        RunRFCheckbox.Locked = False
        MWCheckbox.Locked = False
        SpreadCheckBox.Locked = False
        CreditHoldCheckbox.Locked = False
        DetailTabBox.Locked = False
        CustTabBox.Locked = False
        
        RunRFCheckbox.TabStop = True
        MWCheckbox.TabStop = True
        SpreadCheckBox.TabStop = True
        CreditHoldCheckbox.TabStop = True
        DetailTabBox.TabStop = True
        CustTabBox.TabStop = True
        Run.TabStop = False
        
        RunRFCheckbox.TabIndex = 1
        MWCheckbox.TabIndex = 2
        SpreadCheckBox.TabIndex = 3
        CreditHoldCheckbox.TabIndex = 4
        DetailTabBox.TabIndex = 5
        CustTabBox.TabIndex = 6
        AcctBox.TabIndex = 8
        DocTypeBox.TabIndex = 9
        InvoiceBox.TabIndex = 10
        InvDateBox.TabIndex = 11
        DueDateBox.TabIndex = 12
        OpenAmtBox.TabIndex = 13
        GrossAmtBox.TabIndex = 14
        BUBox.TabIndex = 15
        BU3Box.TabIndex = 16
        BU5Box.TabIndex = 17
        CustAcctBox.TabIndex = 19
        TempCreditBox.TabIndex = 20
        Cancel.TabIndex = 21
        
        Confirm = 0
        AgingForm.Height = AgingForm.Height - 36
        Run.Visible = False
        Cancel.Caption = "RUN REPORT"
    
    ElseIf Confirm = 2 Then
    
        DetailTabBox.BackColor = &H80000005
        CustTabBox.BackColor = &H80000005
        
        For i = 0 To UBound(V)
        
            V(i).BackColor = &H80000005
            V(i).Locked = False
            V(i).TabStop = True
        
        Next i
        
        RunRFCheckbox.Locked = False
        MWCheckbox.Locked = False
        SpreadCheckBox.Locked = False
        CreditHoldCheckbox.Locked = False
        DetailTabBox.Locked = False
        CustTabBox.Locked = False
        
        RunRFCheckbox.TabStop = True
        MWCheckbox.TabStop = True
        SpreadCheckBox.TabStop = True
        CreditHoldCheckbox.TabStop = True
        DetailTabBox.TabStop = True
        CustTabBox.TabStop = True
        Run.TabStop = False
        
        RunRFCheckbox.TabIndex = 1
        MWCheckbox.TabIndex = 2
        SpreadCheckBox.TabIndex = 3
        CreditHoldCheckbox.TabIndex = 4
        DetailTabBox.TabIndex = 5
        CustTabBox.TabIndex = 6
        AcctBox.TabIndex = 8
        DocTypeBox.TabIndex = 9
        InvoiceBox.TabIndex = 10
        InvDateBox.TabIndex = 11
        DueDateBox.TabIndex = 12
        OpenAmtBox.TabIndex = 13
        GrossAmtBox.TabIndex = 14
        BUBox.TabIndex = 15
        BU3Box.TabIndex = 16
        BU5Box.TabIndex = 17
        CustAcctBox.TabIndex = 19
        TempCreditBox.TabIndex = 20
        Cancel.TabIndex = 21
        
        Confirm = 0
        Cancel.Caption = "RUN REPORT"
    End If
    
End Sub

Private Sub CreditHoldCheckbox_Click()

    ToChange = Array(CustTabLabel, CustTabBox, LineLabel, AccountLabel, TempCreditLabel, CustLabel, CustAcctBox, TempCreditBox)
    ToMove = Array(LineLabel2, Cancel, Run)
    
    State = CreditHoldCheckbox.Value
    
    If State = True Then
    
        For i = 0 To UBound(ToChange)
            
            ToChange(i).Visible = True
            
            If TypeName(ToChange(i)) <> "Label" Then
                ToChange(i).Locked = False
            End If
            
        Next i
        
        For i = 0 To UBound(ToMove)
        
            ToMove(i).Top = ToMove(i).Top + 48
            
        Next i
        
        AgingForm.Height = AgingForm.Height + 48
        
    ElseIf State = False Then
    
        For i = 0 To UBound(ToChange)
        
            ToChange(i).Visible = False
            
            If TypeName(ToChange(i)) <> "Label" Then
                ToChange(i).Locked = True
            End If
            
        Next i
        
        For i = 0 To UBound(ToMove)
        
            ToMove(i).Top = ToMove(i).Top - 48
            
        Next i
        
        AgingForm.Height = AgingForm.Height - 48
    
    End If
    
End Sub

Private Sub GotoCust_Click()
    If IsNumeric(CustTabBox.Value) Then
        If (CustTabBox.Value > 0) And (CustTabBox.Value <= Application.Worksheets.Count) Then
            If Sheets(CInt(CustTabBox.Value)).Visible = True Then
                Sheets(CInt(CustTabBox.Value)).Select
                Range("A1").Select
            End If
        End If
    End If
End Sub

Private Sub GotoDetail_Click()
    If IsNumeric(DetailTabBox.Value) Then
        If (DetailTabBox.Value > 0) And (DetailTabBox.Value <= Application.Worksheets.Count) Then
            If Sheets(CInt(DetailTabBox.Value)).Visible = True Then
                Sheets(CInt(DetailTabBox.Value)).Select
                Range("A1").Select
            End If
        End If
    End If
End Sub

Private Sub Run_Click()

    Dim V() As Variant
    
    If Confirm = 1 Then
    
        If CreditHoldCheckbox.Value = True Then
            Q = 11
            QQ = 1
        Else
            Q = 9
            QQ = 0
        End If
        
        V = Array(AcctBox, DocTypeBox, InvoiceBox, InvDateBox, DueDateBox, OpenAmtBox, GrossAmtBox, BUBox, BU3Box, BU5Box, CustAcctBox, TempCreditBox)
        c = ColumnConverter(V)
        t = Array(DetailTabBox, CustTabBox)
        
        Valid = True
        
        For i = 0 To Q
            If c(i) = "!" Then
                Valid = False
            End If
        Next i
        
        For i = 0 To QQ
        
            If IsNumeric(t(i).Value) Then
                If (t(i).Value < 0) Or (t(i).Value > Application.Worksheets.Count) Then
                    Valid = False
                End If
                
                If Sheets(CInt(t(i).Value)).Visible = False Then
                    Valid = False
                End If
            Else
                Valid = False
            End If
            
        Next i
        
        If Valid Then
            RunRF = RunRFCheckbox.Value
            MW = MWCheckbox.Value
            Spread = SpreadCheckBox.Value
            CreditHold = CreditHoldCheckbox.Value
            
            AccountCol = c(0)
            DocTypeCol = c(1)
            InvoiceCol = c(2)
            DateCol = c(3)
            DueCol = c(4)
            OpenCol = c(5)
            DetailTab = t(0).Value
            
            If c(6) = "?" Then
                GrossCol = Array(False, 0)
            Else
                GrossCol = Array(True, c(6))
            End If
            
            If c(7) = "?" Then
                BUCol = Array(False, 0)
            Else
                BUCol = Array(True, c(7))
            End If
            
            If c(8) = "?" Then
                BU3Col = Array(False, 0)
            Else
                BU3Col = Array(True, c(8))
            End If
            
            If c(9) = "?" Then
                BU5Col = Array(False, 0)
            Else
                BU5Col = Array(True, c(9))
            End If
            
            If CreditHold = True Then
                CustAcctCol = c(10)
                TempCreditCol = c(11)
                CustTab = t(1).Value
            Else
                CustAcctCol = 1
                TempCreditCol = 1
                CustTab = 1
            End If
            
            Unload Me
        End If
        
    End If
End Sub

Private Sub Cancel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Debug.Print (KeyAscii)
    
    If KeyAscii = 96 Then
        Tilde = Tilde + 1
    End If
    
End Sub

Private Sub UserForm_Activate()
    Confirm = 0
    AgingForm.Height = AgingForm.Height - 36
    Run.Visible = False
    Cancel.Caption = "RUN REPORT"
    Tilde = 0
End Sub

Private Sub UserForm_Initialize()
    
    Tilde = 0
    
    Run.Visible = False
    Cancel.Caption = "RUN REPORT"
    Confirm = 0
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        RunRF = False
        RunNSF = False
        CreditHold = False
        Under10 = False
        Spread = False
        MW = False
    End If
    
    AcctBox.Value = Null
    DocTypeBox.Value = Null
    InvoiceBox.Value = Null
    InvDateBox.Value = Null
    DueDateBox.Value = Null
    OpenAmtBox.Value = Null
    GrossAmtBox.Value = Null
    BUBox.Value = Null
    BU3Box.Value = Null
    BU5Box.Value = Null
    AcctBox.BackColor = &H80000005
    DocTypeBox.BackColor = &H80000005
    InvoiceBox.BackColor = &H80000005
    InvDateBox.BackColor = &H80000005
    DueDateBox.BackColor = &H80000005
    OpenAmtBox.BackColor = &H80000005
    GrossAmtBox.BackColor = &H80000005
    BUBox.BackColor = &H80000005
    BU3Box.BackColor = &H80000005
    BU5Box.BackColor = &H80000005
    
End Sub

Private Function ColumnConverter(Inputs() As Variant) As Variant

    Dim Data() As Variant
    Data = Inputs
    ReDim Res(0 To UBound(Data)) As Variant
    
    For i = 0 To UBound(Data)
    
        If IsNumeric(Data(i)) And (Data(i) > 0) Then
            Res(i) = WorksheetFunction.Min(CInt(Data(i)), 16384)
        ElseIf IsLetter(Data(i)) And (Len(Data(i)) < 4) Then
            Res(i) = Range(UCase(Data(i)) & 1).Column
        Else
            Res(i) = "!"
        End If
        
        If (Data(i).Value = "") And (Data(i).Tag = "0") Then
            Res(i) = "?"
        End If
    
    Next i
    
    ColumnConverter = Res

End Function

'Private Function ColumnConverter(Acct As Variant, Doc As Variant, Inv As Variant, Dat As Variant, _
'    Due As Variant, Op As Variant, CustAcct As Variant, TempCredit As Variant, Gross As Variant, BU As Variant, BU3 As Variant, BU5 As Variant)
'
'    If IsNumeric(Acct) And (Acct > 0) Then
'        Account = WorksheetFunction.Min(CInt(Acct), 16384)
'    ElseIf IsLetter(Acct) And (Len(Acct) < 4) Then
'        Account = Range(UCase(Acct) & 1).Column
'    Else
'        Account = "!"
'    End If
'
'    If IsNumeric(Doc) And (Doc > 0) Then
'        DocType = WorksheetFunction.Min(CInt(Doc), 16384)
'    ElseIf IsLetter(Doc) And (Len(Doc) < 4) Then
'        DocType = Range(UCase(Doc) & 1).Column
'    Else
'        DocType = "!"
'    End If
'
'    If IsNumeric(Inv) And (Inv > 0) Then
'        Invoice = WorksheetFunction.Min(CInt(Inv), 16384)
'    ElseIf IsLetter(Inv) And (Len(Inv) < 4) Then
'        Invoice = Range(UCase(Inv) & 1).Column
'    Else
'        Invoice = "!"
'    End If
'
'    If IsNumeric(Dat) And (Dat > 0) Then
'        InvDate = WorksheetFunction.Min(CInt(Dat), 16384)
'    ElseIf IsLetter(Dat) And (Len(Dat) < 4) Then
'        InvDate = Range(UCase(Dat) & 1).Column
'    Else
'        InvDate = "!"
'    End If
'
'    If IsNumeric(Due) And (Due > 0) Then
'        DueDate = WorksheetFunction.Min(CInt(Due), 16384)
'    ElseIf IsLetter(Due) And (Len(Due) < 4) Then
'        DueDate = Range(UCase(Due) & 1).Column
'    Else
'        DueDate = "!"
'    End If
'
'    If IsNumeric(Op) And (Op > 0) Then
'        Openn = WorksheetFunction.Min(CInt(Op), 16384)
'    ElseIf IsLetter(Op) And (Len(Op) < 4) Then
'        Openn = Range(UCase(Op) & 1).Column
'    Else
'        Openn = "!"
'    End If
'
'    If IsNumeric(CustAcct) And (CustAcct > 0) Then
'        CstAcct = WorksheetFunction.Min(CInt(CustAcct), 16384)
'    ElseIf IsLetter(CustAcct) And (Len(CustAcct) < 4) Then
'        CstAcct = Range(UCase(CustAcct) & 1).Column
'    Else
'        CstAcct = "!"
'    End If
'
'    If IsNumeric(TempCredit) And (TempCredit > 0) Then
'        TC = WorksheetFunction.Min(CInt(TempCredit), 16384)
'    ElseIf IsLetter(TempCredit) And (Len(TempCredit) < 4) Then
'        TC = Range(UCase(TempCredit) & 1).Column
'    Else
'        TC = "!"
'    End If
'
'    If IsNumeric(Gross) And (Gross > 0) Then
'        GrossAmt = WorksheetFunction.Min(CInt(Gross), 16384)
'    ElseIf IsLetter(Gross) And (Len(Gross) < 4) Then
'        GrossAmt = Range(UCase(Gross) & 1).Column
'    Else
'        GrossAmt = "!"
'    End If
'
'    If IsNumeric(BU) And (BU > 0) Then
'        BusUnit = WorksheetFunction.Min(CInt(BU), 16384)
'    ElseIf IsLetter(Gross) And (Len(Gross) < 4) Then
'        GrossAmt = Range(UCase(Gross) & 1).Column
'    Else
'        GrossAmt = "!"
'    End If
'
'    If IsNumeric(Gross) And (Gross > 0) Then
'        GrossAmt = WorksheetFunction.Min(CInt(Gross), 16384)
'    ElseIf IsLetter(Gross) And (Len(Gross) < 4) Then
'        GrossAmt = Range(UCase(Gross) & 1).Column
'    Else
'        GrossAmt = "!"
'    End If
'
'    If IsNumeric(Gross) And (Gross > 0) Then
'        GrossAmt = WorksheetFunction.Min(CInt(Gross), 16384)
'    ElseIf IsLetter(Gross) And (Len(Gross) < 4) Then
'        GrossAmt = Range(UCase(Gross) & 1).Column
'    Else
'        GrossAmt = "!"
'    End If
'
'    ColumnConverter = Array(Account, DocType, Invoice, InvDate, DueDate, Openn, CstAcct, TC)
'
'End Function

Private Function IsLetter(Str)

    For i = 1 To Len(Str)
        
        Select Case Asc(Mid(Str, i, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function
