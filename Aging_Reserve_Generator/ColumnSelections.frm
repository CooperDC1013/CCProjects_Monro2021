VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColumnSelections 
   Caption         =   "Select Aging Columns"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "ColumnSelections.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColumnSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private State As Integer

Private Function ResetParams()

    Cont = Array(TabB, TypeB, AddressB, AlphaB, OpenB, CurrB, B30, B60, B90, B120, B150)
    
    For i = 0 To UBound(Cont)
        
        Cont(i).Value = Null
        Cont(i).BackColor = &H8000000F
        Cont(i).Locked = False
        
    Next i
    
    TabB.SetFocus
    
    Launch.Caption = "NOT READY"
    Launch.BackColor = &H8080FF
    Validate.Caption = "VALIDATE"
    
    State = 0
        
End Function

Private Function ColumnValidate(Entry As String) As Integer
    
    If IsNumeric(Entry) Then
    
        If CInt(Entry) <= 16384 And CInt(Entry) > 0 Then
            ColumnValidate = CInt(Entry)
        Else
            ColumnValidate = 0
        End If
    
    ElseIf IsLetter(Entry) And Len(Entry) < 4 Then
        ColumnValidate = CInt(Columns(Entry).Column)
    Else
        ColumnValidate = 0
    End If

End Function

Private Function IsLetter(Str) As Boolean

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

Private Sub GotoTab_Click()
    If IsNumeric(TabB.Value) Then
        If CInt(TabB.Value) <= Application.Worksheets.Count And CInt(TabB.Value) > 0 Then
            If Sheets(CInt(TabB.Value)).Visible = True Then
                Sheets(CInt(TabB.Value)).Select
                Range("A1").Select
            End If
        End If
    End If
End Sub


Private Sub Reset_Click()
    ResetParams
End Sub

Private Sub UserForm_Activate()
    ResetParams
End Sub

Private Sub UserForm_Initialize()
    ResetParams
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Reserve.Abort = False
    
    If CloseMode <> 1 Then
        Reserve.Abort = True
    End If
    
    ResetParams
    
End Sub

Private Sub Validate_Click()

    Cont = Array(TypeB, AddressB, AlphaB, OpenB, CurrB, B30, B60, B90, B120, B150)
    
    If State = 0 Then

        Valid = False
        ValidControls = True
        ValidTab = False
        
        For i = 0 To UBound(Cont)
        
            Cont(i).Locked = True
            
            If IsLetter(Cont(i).Value) Then
                Cont(i).Value = UCase(Cont(i).Value)
            End If
            
        Next i
        
        For i = 0 To UBound(Cont)
            
            If ColumnValidate(Cont(i).Value) = 0 Then
                ValidControls = False
                Cont(i).BackColor = &H8080FF
            Else
                Cont(i).BackColor = &H80FF80
            End If
        
        Next i
        
        TabB.Locked = True
        
        TabBI = IsNumeric(TabB.Value)
        
        If TabBI Then
            If TabB.Value > 0 And TabB.Value <= Application.Worksheets.Count Then
                TabBII = True
            Else
                TabBII = False
            End If
        Else
            TabBII = False
        End If
        
        If TabBII Then
            If Sheets(CInt(TabB.Value)).Visible = True Then
                TabBIII = True
            Else
                TabBIII = False
            End If
        Else
            TabBIII = False
        End If
        
        If TabBI = False Or TabBII = False Or TabBIII = False Then
            ValidTab = False
            TabB.BackColor = &H8080FF
        Else
            ValidTab = True
            TabB.BackColor = &H80FF80
        End If
        
        If ValidTab And ValidControls Then
            Valid = True
        Else
            Valid = False
        End If
        
        If Valid Then
            
            State = 1
            
            Launch.Caption = "LAUNCH"
            Launch.BackColor = &H80FF80
            Launch.SetFocus
            
        Else
            
            State = -1
            
            Launch.Caption = "NOT READY"
            Launch.BackColor = &H8080FF
            
        End If
        
        Validate.Caption = "EDIT"
        
    ElseIf State = -1 Then
        
        State = 0
        
        Validate.Caption = "VALIDATE"
        Launch.Caption = "NOT READY"
        Launch.BackColor = &H8080FF
        
        For i = 0 To UBound(Cont)
        
            Cont(i).BackColor = &H80000005
            Cont(i).Locked = False
            
        Next i
        
        TabB.BackColor = &H80000005
        TabB.Locked = False
        
    ElseIf State = 1 Then
    
        State = 0
    
        Validate.Caption = "VALIDATE"
        Launch.Caption = "NOT READY"
        Launch.BackColor = &H8080FF
        
        For i = 0 To UBound(Cont)
        
            Cont(i).BackColor = &H80000005
            Cont(i).Locked = False
            
        Next i
        
        TabB.BackColor = &H80000005
        TabB.Locked = False
        
    End If
    
End Sub

Private Sub Launch_Click()

    If State = 1 Then
        
        Aging = ColumnValidate(TabB.Value)
        CType = ColumnValidate(TypeB.Value)
        Address = ColumnValidate(AddressB.Value)
        Alpha = ColumnValidate(AlphaB.Value)
        COpen = ColumnValidate(OpenB.Value)
        Curr = ColumnValidate(CurrB.Value)
        C30 = ColumnValidate(B30.Value)
        C60 = ColumnValidate(B60.Value)
        C90 = ColumnValidate(B90.Value)
        C120 = ColumnValidate(B120.Value)
        C150 = ColumnValidate(B150.Value)
        
        Unload Me
    
    End If
    
End Sub
