VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalTranForm2 
   Caption         =   "CalTrans Template"
   ClientHeight    =   9870.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   OleObjectBlob   =   "CalTranForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CalTranForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Brand As String
Private StartD As Variant
Private FinishD As Variant
Private Head As String
Private StartValid As Boolean
Private FinishValid As Boolean
Private State As Integer
Private Col As Variant
Private Buttons As Variant

Private Sub AllenTireL_Click()
 Update
End Sub

Private Sub ARButton_Click()
    If IsNumeric(ARTab.Value) And ARTab.Value > 0 And ARTab.Value <= Application.Worksheets.Count Then
        If Sheets(CInt(ARTab.Value)).Visible = True Then
            Application.Worksheets(CInt(ARTab.Value)).Select
            Range("A1").Select
        End If
    End If
End Sub

Private Sub ARHeader_Click()
    Update
End Sub

Private Sub ARInvCol_Change()
    If IsNumeric(ARTab.Value) And ARTab.Value > 0 And ARTab.Value <= Application.Worksheets.Count Then
        If Sheets(CInt(ARTab.Value)).Visible = True Then
            Application.Worksheets(CInt(ARTab.Value)).Select
        End If
    End If
End Sub

Private Sub Help_Click()
    HelpScreen.Show vbModal
End Sub

Private Sub Update()

    Headers = 0

    Buttons = Array(AllenTireL, MONROL, TireChoiceL, MrTireL, TiresNowL, Vacant1, Vacant2, Vacant3, NONE)
    BrandT = Array("Allen Tire", "MONRO", "Tire Choice", "Mr. Tire", "Tires Now", "No Brand", "No Brand", "No Brand", "No Brand")
    
    For i = 0 To 8
        If Buttons(i).Value = True Then
            Brand = BrandT(i)
        End If
    Next i
    
    If (ARHeader.Value And STOREHeader.Value) Then
        Headers = 3
    ElseIf STOREHeader.Value = True Then
        Headers = 2
    ElseIf ARHeader.Value = True Then
        Headers = 1
    Else
        Headers = 0
    End If
    
    Select Case Headers
        Case 0
            Head = "No Headers"
        Case 1
            Head = "AR Only"
        Case 2
            Head = "STORE Only"
        Case 3
            Head = "Both Tabs"
    End Select
    
    Selections.Text = _
        ":: Selections ::" & vbCr & vbCr & _
        " Brand  : " & Brand & vbCr & vbCr & _
        " Begin  : " & StartD & vbCr & _
        "  End   : " & FinishD & vbCr & vbCr & _
        "Headers : " & Head

End Sub

Private Sub MONROL_Click()
    Update
End Sub

Private Sub MrTireL_Click()
    Update
End Sub

Private Sub NONE_Click()
    Update
End Sub

Private Sub Reset_Click()
    
    If State = 0 Then
    
        Run.Caption = "CONFIRM"
        Reset.Caption = "RESET SELECTIONS"
        Selections.BackColor = &H8000000F
        StartSet.Locked = False
        FinishSet.Locked = False
        ARHeader.Locked = False
        STOREHeader.Locked = False
        
        Help.Locked = False
        Run.Locked = False
        
        For i = 0 To UBound(Buttons)
            Buttons(i).Locked = False
            Buttons(i).Value = False
        Next i
        
        For Each Control In CalTranForm2.Controls
        
            If Control.Name = "Reset" Or Control.Name = "Selections" Then
                Control.TabStop = False
            ElseIf TypeName(Control) = "Label" Then
            
            Else
                Control.TabStop = True
            End If
        
        Next Control
        
        For i = 0 To UBound(Col)
        
            Col(i).Value = Null
            Col(i).BackColor = &H80000005
            Col(i).Locked = False
            
        Next i
        
        StartDate.Value = Null
        EndDate.Value = Null
        StartDate.BackColor = &H80000005
        EndDate.BackColor = &H80000005
        
        Update
        
    ElseIf State = 1 Then
    
        State = 0
    
        Run.Caption = "CONFIRM"
        Reset.Caption = "RESET SELECTIONS"
        Selections.BackColor = &H8000000F
        StartSet.Locked = False
        FinishSet.Locked = False
        ARHeader.Locked = False
        STOREHeader.Locked = False
        
        Help.Locked = False
        
        For i = 0 To UBound(Buttons)
            Buttons(i).Locked = False
        Next i
        
        For Each Control In CalTranForm2.Controls
        
            If Control.Name = "Reset" Or Control.Name = "Selections" Then
                Control.TabStop = False
            ElseIf TypeName(Control) = "Label" Then
            
            Else
                Control.TabStop = True
            End If
        
        Next Control
        
        For i = 0 To UBound(Col)
        
            Col(i).BackColor = &H80000005
            Col(i).Locked = False
            
        Next i
        
        StartDate.BackColor = &H80000005
        EndDate.BackColor = &H80000005
        
        Update
        
    ElseIf State = 2 Then
    
        Run.Caption = "CONFIRM"
        Reset.Caption = "RESET SELECTIONS"
        Selections.BackColor = &H8000000F
        StartSet.Locked = False
        FinishSet.Locked = False
        ARHeader.Locked = False
        STOREHeader.Locked = False
        
        Help.Locked = False
        Run.Locked = False
        
        For i = 0 To UBound(Buttons)
            Buttons(i).Locked = False
        Next i
        
        For Each Control In CalTranForm2.Controls
        
            If Control.Name = "Reset" Or Control.Name = "Selections" Then
                Control.TabStop = False
            ElseIf TypeName(Control) = "Label" Then
            
            Else
                Control.TabStop = True
            End If
        
        Next Control
        
        For i = 0 To UBound(Col)
        
            Col(i).BackColor = &H80000005
            Col(i).Locked = False
            
        Next i
        
        StartDate.BackColor = &H80000005
        EndDate.BackColor = &H80000005
        
        State = 0
        
        Update
        
    End If
        
End Sub

Private Sub Run_Click()
    
    Res = ColumnConverter(Col)
    
    If State = 0 Then
    
        Valid = True
        
        Res = ColumnConverter(Col)
        
        StartSet_Click
        FinishSet_Click
        
        If StartD = "INVALID" Or FinishD = "INVALID" Then
            Valid = False
        End If
        
        If StartD = "" Or FinishD = "" Then
            Valid = False
        End If
        
        If Not IsNumeric(ARTab.Value) And ARTab.Value > 0 And ARTab.Value <= Application.Worksheets.Count Then
            If Sheets(CInt(ARTab.Value)).Visible = True Then
                Valid = False
                Res(0) = "!"
            End If
        End If

        If Not IsNumeric(StoreTab.Value) And StoreTab.Value > 0 And StoreTab.Value <= Application.Worksheets.Count Then
            If Sheets(CInt(StoreTab.Value)).Visible = True Then
                Valid = False
                Res(1) = "!"
            End If
        End If
        
        For i = 0 To UBound(Res)
        
            If Res(i) = "!" Then
                Valid = False
                Col(i).BackColor = &HFF&
            ElseIf Res(i) <> "!" Then
                Col(i).BackColor = &HFF00&
            End If
        
        Next i
        
        ButtonPos = False
        
        For i = 0 To UBound(Buttons)
            If Buttons(i).Value = True Then
                ButtonPos = True
                Exit For
            End If
        Next i
        
        If ButtonPos = False Then
            NONE.Value = True
        End If
        
        If Valid Then
            State = 1 'Valid State
            Run.Caption = "RUN"
            Reset.Caption = "EDIT SELECTIONS"
            Selections.BackColor = &HFF00&
            StartSet.Locked = True
            FinishSet.Locked = True
            ARHeader.Locked = True
            STOREHeader.Locked = True
            
            Help.Locked = True
            
            For i = 0 To UBound(Col)
                Col(i).Locked = True
            Next i
            
            For i = 0 To UBound(Buttons)
                Buttons(i).Locked = True
            Next i
            
            For Each Control In CalTranForm2.Controls
            
                If Control.Name = "Reset" Or Control.Name = "Run" Then
                    Control.TabStop = True
                ElseIf TypeName(Control) = "Label" Then
                
                Else
                    Control.TabStop = False
                End If
            
            Next Control
        
        Else
            State = 2 'Error State
            Run.Caption = ""
            Run.Locked = True
            Reset.Caption = "EDIT SELECTIONS"
            Selections.BackColor = &HFF&
            StartSet.Locked = True
            FinishSet.Locked = True
            ARHeader.Locked = True
            STOREHeader.Locked = True
            
            For i = 0 To UBound(Col)
                Col(i).Locked = True
            Next i
            
            For i = 0 To UBound(Buttons)
                Buttons(i).Locked = True
            Next i
            
            For Each Control In CalTranForm2.Controls
            
                If Control.Name = "Reset" Then
                    Control.TabStop = True
                ElseIf TypeName(Control) = "Label" Then
                
                Else
                    Control.TabStop = False
                End If
            
            Next Control
            
        End If
        
    ElseIf State = 1 Then
    
        State = 0
    
        If ARHeader.Value Then ARHead = 2 Else ARHead = 1
        If STOREHeader.Value Then StoreHead = 2 Else StoreHead = 1
        
        AR = Res(0)
        Store = Res(1)
        ARInv = Res(2)
        ARDate = Res(3)
        StoreN = Res(4)
        StInv = Res(5)
        Gross = Res(6)
        PO = Res(7)
        Tax = Res(8)
        Taxable = Res(9)
        Make = Res(10)
        Model = Res(11)
        Mile = Res(12)
        Lic = Res(13)
        VIN = Res(14)
        ItemN = Res(15)
        Desc = Res(16)
        Writer = Res(17)
        Parts = Res(18)
        Labor = Res(19)
        Qty = Res(20)
        
        Start = StartD
        Finish = FinishD
        
        AllenTire = Buttons(0).Value
        Monro = Buttons(1).Value
        TireChoice = Buttons(2).Value
        MrTire = Buttons(3).Value
        TiresNow = Buttons(4).Value
        Vac1 = Buttons(5).Value
        Vac2 = Buttons(6).Value
        Vac3 = Buttons(7).Value
        
        Confirm = True
        
        Unload Me
    End If
    
End Sub

Private Sub StartSet_Click()
    If IsDate(CStr(StartDate.Value)) Then
        If CDate(StartDate.Value) <= FinishD Or IsDate(FinishD) = False Then
            StartD = CDate(StartDate.Value)
            StartValid = True
            StartDate.BackColor = &H80000005
            Update
        Else
            StartValid = False
            StartD = "INVALID"
            StartDate.BackColor = &HFF&
            Update
        End If
    ElseIf StartDate.Value = "" Then
        StartD = ""
        StartValid = False
        StartDate.BackColor = &H80000005
        Update
    Else
        StartValid = False
        StartD = "INVALID"
        StartDate.BackColor = &HFF&
        Update
    End If
End Sub

Private Sub FinishSet_Click()
    If IsDate(CStr(EndDate.Value)) Then
        If CDate(EndDate.Value) >= StartD Or IsDate(StartD) = False Then
            FinishD = CDate(EndDate.Value)
            FinishValid = True
            EndDate.BackColor = &H80000005
            Update
        Else
            FinishValid = False
            FinishD = "INVALID"
            EndDate.BackColor = &HFF&
            Update
        End If
    ElseIf EndDate.Value = "" Then
        FinishD = ""
        FinishValid = False
        EndDate.BackColor = &H80000005
        Update
    Else
        FinishValid = False
        FinishD = "INVALID"
        EndDate.BackColor = &HFF&
        Update
    End If
End Sub

Private Sub StartToday_Click()
    StartDate.Value = ""
    StartDate.BackColor = &H80000005
    StartDate.Value = (Date - 1)
End Sub

Private Sub FinishToday_Click()
    EndDate.Value = ""
    EndDate.BackColor = &H80000005
    EndDate.Value = (Date - 1)
End Sub

Private Sub StoreButton_Click()
    If IsNumeric(StoreTab.Value) And StoreTab.Value > 0 And StoreTab.Value <= Application.Worksheets.Count Then
        If Sheets(CInt(StoreTab.Value)).Visible = True Then
            Application.Worksheets(CInt(StoreTab.Value)).Select
            Range("A1").Select
        End If
    End If
End Sub

Private Sub StoreCol_Change()
    If IsNumeric(StoreTab.Value) And StoreTab.Value > 0 And StoreTab.Value <= Application.Worksheets.Count Then
        If Sheets(CInt(StoreTab.Value)).Visible = True Then
            Application.Worksheets(CInt(StoreTab.Value)).Select
        End If
    End If
End Sub

Private Sub STOREHeader_Click()
    Update
End Sub

Private Sub TireChoiceL_Click()
    Update
End Sub

Private Sub TiresNowL_Click()
    Update
End Sub

Private Sub UserForm_Activate()
    State = 0
    StartValid = False
    FinishValid = False
    StartDate.Value = ""
    EndDate.Value = ""
    AllenTireL.Value = False
    MONROL.Value = False
    TireChoiceL.Value = False
    MrTireL.Value = False
    TiresNowL.Value = False
    NONE.Value = False
    
    For Each Control In CalTranForm2.Controls
            
        If Control.Name = "Reset" Or Control.Name = "Selections" Then
            Control.TabStop = False
        ElseIf TypeName(Control) = "Label" Then
        
        Else
            Control.TabStop = True
        End If
            
    Next Control
    
     Col = Array(ARTab, StoreTab, ARInvCol, ARInvDateCol, StoreCol, STOREInvCol, GrossCol, POCol, TaxCol, _
        TaxableCol, MakeCol, ModelCol, MileageCol, LicCol, VINCol, ItemCol, ItemDescCol, ServiceCol, PartsCol, _
        LaborCol, QtyCol)
        
    Buttons = Array(AllenTireL, MONROL, TireChoiceL, MrTireL, TiresNowL, Vacant1, Vacant2, Vacant3, NONE)
    
    Update
End Sub

Private Sub UserForm_Initialize()
    State = 0
    StartValid = False
    FinishValid = False
    StartDate.Value = ""
    EndDate.Value = ""
    AllenTireL.Value = False
    MONROL.Value = False
    TireChoiceL.Value = False
    MrTireL.Value = False
    TiresNowL.Value = False
    NONE.Value = False
    
    For Each Control In CalTranForm2.Controls
            
        If Control.Name = "Reset" Or Control.Name = "Selections" Then
            Control.TabStop = False
        ElseIf TypeName(Control) = "Label" Then
            
        Else
            Control.TabStop = True
        End If
            
    Next Control
    
     Col = Array(ARTab, StoreTab, ARInvCol, ARInvDateCol, StoreCol, STOREInvCol, GrossCol, POCol, TaxCol, _
        TaxableCol, MakeCol, ModelCol, MileageCol, LicCol, VINCol, ItemCol, ItemDescCol, ServiceCol, PartsCol, _
        LaborCol, QtyCol)
        
    Buttons = Array(AllenTireL, MONROL, TireChoiceL, MrTireL, TiresNowL, Vacant1, Vacant2, Vacant3, NONE)
    
    Update
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Confirm = False
    End If
    
    State = 0
    
    Run.Caption = "CONFIRM"
    Reset.Caption = "RESET SELECTIONS"
    Selections.BackColor = &H8000000F
    StartSet.Locked = False
    FinishSet.Locked = False
    ARHeader.Locked = False
    STOREHeader.Locked = False
    
    Help.Locked = False
    
    For i = 0 To UBound(Buttons)
        Buttons(i).Locked = False
    Next i
    
    For Each Control In CalTranForm2.Controls
    
        If Control.Name = "Reset" Or Control.Name = "Selections" Then
            Control.TabStop = False
        ElseIf TypeName(Control) = "Label" Then
        
        Else
            Control.TabStop = True
        End If
    
    Next Control
    
    For i = 0 To UBound(Col)
    
        Col(i).Value = Null
        Col(i).BackColor = &H80000005
        Col(i).Locked = False
        
    Next i
    
    StartDate.Value = Null
    EndDate.Value = Null
    StartDate.BackColor = &H80000005
    EndDate.BackColor = &H80000005
    
End Sub

Private Sub Vacant1_Click()
    Update
End Sub

Private Sub Vacant2_Click()
    Update
End Sub

Private Sub Vacant3_Click()
    Update
End Sub

Private Function ColumnConverter(Values As Variant) As Variant

    Dim Res(0 To 20) As Variant
    
    For i = 0 To UBound(Values)
    
        If IsNumeric(Values(i).Value) And (Values(i).Value > 0) Then
            Res(i) = WorksheetFunction.Min(CInt(Values(i).Value), 16384)
        ElseIf IsLetter(Values(i).Value) And (Len(Values(i).Value) < 4) Then
            Res(i) = Range(UCase(Values(i).Value) & 1).Column
        Else
            Res(i) = "!"
        End If
    
    Next i
    
    ColumnConverter = Res

End Function

Private Function IsLetter(Str)

    For i = 1 To Len(Str)
    
        Select Case Asc(Mid(Str, i, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
        
    Next i

End Function
