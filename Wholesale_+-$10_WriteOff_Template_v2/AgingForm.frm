VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AgingForm 
   Caption         =   "Launch Program"
   ClientHeight    =   9300.001
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
        RunText = "Nothin' to see here, turd."
    Else
        RunText = "CONFIRM RUN"
    End If
    
    Tilde = 0
    
    V = Array(AcctBox, InvoiceBox, InvDateBox, BUBox, BU3Box, BU5Box, DocTypeBox, GrossAmtBox, OpenAmtBox)
    t = Array(DetailTabBox)
    
    If Confirm = 0 Then
    
        C = ColumnConverter(V)
        
        Valid = True
        
        For i = 0 To UBound(C)
            If C(i) = "!" Then
                Valid = False
            End If
        Next i
        
        For i = 0 To UBound(t)
        
            If IsNumeric(t(i).Value) And (t(i).Value > 0) And (t(i).Value <= Application.Worksheets.Count) Then
                If Sheets(CInt(t(i).Value)).Visible = False Then
                    Valid = False
                End If
            Else
                Valid = False
            End If
            
        Next i
        
        If Not IsDate(CStr(GLDate.Value)) Then
            GLDate.BackColor = &HFF
            Valid = False
        End If
        
        If Valid Then
        
            DetailTabBox.BackColor = &H80FF80
            GLDate.BackColor = &H80FF80
            GLDate.Locked = True
            GLDate.TabStop = False
            SpreadCheckBox.Locked = True
            DetailTabBox.Locked = True
            SpreadCheckBox.TabStop = False
            DetailTabBox.TabStop = False
            
            For i = 0 To UBound(C)
            
                V(i).BackColor = &H80FF80
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
                If C(i) = "!" Then
                    V(i).BackColor = &HFF
                Else
                    V(i).BackColor = &H80FF80
                End If
                
                V(i).Locked = True
                V(i).TabStop = False
            Next i
            
            SpreadCheckBox.Locked = True
            SpreadCheckBox.TabStop = False
            
            DetailTabBox.TabStop = False
            DetailTabBox.Locked = True
            Cancel.TabIndex = 0
            
            GLDate.TabStop = False
            GLDate.Locked = True
            
            Confirm = 2
            Cancel.Caption = "RETURN"
            
            For i = 0 To UBound(t)
        
                If IsNumeric(t(i).Value) And (t(i).Value > 0) And (t(i).Value <= Application.Worksheets.Count) Then
                
                    t(i).BackColor = &H80FF80
                    
                    If Sheets(CInt(t(i).Value)).Visible = False Then
                        t(i).BackColor = &HFF
                    End If
                Else
                    t(i).BackColor = &HFF
                End If
                
                t(i).Locked = True
                
            Next i
            
        End If
        
    ElseIf Confirm = 1 Then
    
        DetailTabBox.BackColor = &H80000005
        GLDate.BackColor = &H80000005
                
        For i = 0 To UBound(V)
        
            V(i).BackColor = &H80000005
            V(i).Locked = False
            V(i).TabStop = True
        
        Next i
        
        SpreadCheckBox.Locked = False
        DetailTabBox.Locked = False
        SpreadCheckBox.TabStop = True
        DetailTabBox.TabStop = True
        GLDate.Locked = False
        GLDate.TabStop = True
        
        Run.TabStop = False
        
        SpreadCheckBox.TabIndex = 1
        DetailTabBox.TabIndex = 2
        GotoDetail.TabIndex = 3
        AcctBox.TabIndex = 4
        InvoiceBox.TabIndex = 5
        InvDateBox.TabIndex = 6
        BUBox.TabIndex = 7
        BU3Box.TabIndex = 8
        BU5Box.TabIndex = 9
        DocTypeBox.TabIndex = 10
        GrossAmtBox.TabIndex = 11
        OpenAmtBox.TabIndex = 12
        GLDate.TabIndex = 13
        Cancel.TabIndex = 14
        
        Confirm = 0
        AgingForm.Height = AgingForm.Height - 36
        Run.Visible = False
        Cancel.Caption = "RUN REPORT"
    
    ElseIf Confirm = 2 Then
    
        DetailTabBox.BackColor = &H80000005
        GLDate.BackColor = &H80000005
        
        For i = 0 To UBound(V)
        
            V(i).BackColor = &H80000005
            V(i).Locked = False
            V(i).TabStop = True
        
        Next i
        
        SpreadCheckBox.Locked = False
        DetailTabBox.Locked = False
        SpreadCheckBox.TabStop = True
        DetailTabBox.TabStop = True
        GLDate.TabStop = True
        GLDate.Locked = False
        
        Run.TabStop = False
        
        SpreadCheckBox.TabIndex = 1
        DetailTabBox.TabIndex = 2
        GotoDetail.TabIndex = 3
        AcctBox.TabIndex = 4
        InvoiceBox.TabIndex = 5
        InvDateBox.TabIndex = 6
        BUBox.TabIndex = 7
        BU3Box.TabIndex = 8
        BU5Box.TabIndex = 9
        DocTypeBox.TabIndex = 10
        GrossAmtBox.TabIndex = 11
        OpenAmtBox.TabIndex = 12
        GLDate.TabIndex = 13
        Cancel.TabIndex = 14
        
        Confirm = 0
        Cancel.Caption = "RUN REPORT"
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

Private Sub Label16_Click()

End Sub

Private Sub Run_Click()

    Dim V() As Variant
    
    If Confirm = 1 Then
        
        V = Array(AcctBox, InvoiceBox, InvDateBox, BUBox, BU3Box, BU5Box, DocTypeBox, GrossAmtBox, OpenAmtBox)
        C = ColumnConverter(V)
        t = Array(DetailTabBox)
        
        Valid = True
        
        For i = 0 To UBound(C)
            If C(i) = "!" Then
                Valid = False
            End If
        Next i
        
        For i = 0 To UBound(t)
        
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
        
        If Not IsDate(CStr(GLDate.Value)) Then
            Valid = False
        End If
        
        If Valid Then
        
            Spread = SpreadCheckBox.Value
            
            AccountCol = C(0)
            InvoiceCol = C(1)
            DateCol = C(2)
            BUCol = C(3)
            BU3Col = C(4)
            BU5Col = C(5)
            DocTypeCol = C(6)
            GrossCol = C(7)
            OpenCol = C(8)
            DetailTab = t(0).Value
            
            GLD = CDate(CStr(GLDate.Value))
            
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
        Spread = False
    End If
    
    AcctBox.Value = Null
    InvoiceBox.Value = Null
    InvDateBox.Value = Null
    BUBox.Value = Null
    BU3Box.Value = Null
    BU5Box.Value = Null
    DocTypeBox.Value = Null
    GrossAmtBox.Value = Null
    OpenAmtBox.Value = Null
    GLDate.Value = Null
    
    AcctBox.BackColor = &H80000005
    InvoiceBox.BackColor = &H80000005
    InvDateBox.BackColor = &H80000005
    BUBox.BackColor = &H80000005
    BU3Box.BackColor = &H80000005
    BU5Box.BackColor = &H80000005
    DocTypeBox.BackColor = &H80000005
    GrossAmtBox.BackColor = &H80000005
    OpenAmtBox.BackColor = &H80000005
    GLDate.BackColor = &H80000005
    
    DetailTabBox.BackColor = &H80000005
    DetailTabBox.Value = Null
    
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
