VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_PaymentTool 
   Caption         =   "Launch Payment Tool"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "F_PaymentTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_PaymentTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private State As Integer

Private Sub BalOpt_Click()
    F_PaymentTool.Height = 426
    
    AsOfLabel.Visible = False
    AsOfVal.Visible = False
    
    LBalVal.Visible = True
    BalLabel.Visible = True
    UBalVal.Visible = True
End Sub

Private Sub DateOpt_Click()
    F_PaymentTool.Height = 426
    
    AsOfLabel.Visible = True
    AsOfVal.Visible = True
    
    LBalVal.Visible = False
    BalLabel.Visible = False
    UBalVal.Visible = False
    
End Sub

Private Sub GotoTab_Click()
    If TabLoc.Value > 0 And TabLoc.Value < Application.Worksheets.Count Then
        Application.Sheets(CInt(TabLoc.Value)).Select
        Range("A1").Select
    End If
End Sub

Private Sub Reset_Click()
    ResetParams
    F_PaymentTool.Height = 146
End Sub

Private Sub Run_Click()
    
    Cols = Array(RPDIVJV, RPDOCV, RPAGV, RPDCTV, RPDCTMV, RPGLBAV, RPDGJV)
    
    Select Case State
        
        Case 0
            
            Valid = True
            
            AsOfVal.Locked = True
            LBalVal.Locked = True
            UBalVal.Locked = True
            TabLoc.Locked = True
            DateOpt.Locked = True
            BalOpt.Locked = True
            
            For Each Entry In Cols
            
                Entry.BackColor = &H80FF80
                Entry.Locked = True
                
                If ColumnValidate(Entry.Value) = False Then
                    Valid = False
                    Entry.BackColor = &H8080FF
                End If
                
            Next Entry
            
            If IsNumeric(TabLoc.Value) Then
                If TabLoc.Value < 0 Or TabLoc.Value > Application.Worksheets.Count Then
                    Valid = False
                    TabLoc.BackColor = &H8080FF
                Else
                    TabLoc.BackColor = &H80FF80
                End If
            Else
                Valid = False
                TabLoc.BackColor = &H8080FF
            End If
            
            If (BalRangeValidate(LBalVal.Value, UBalVal.Value) = False) And BalOpt.Value = True Then
                Valid = False
                LBalVal.BackColor = &H8080FF
                UBalVal.BackColor = &H8080FF
            Else
                LBalVal.BackColor = &H80FF80
                UBalVal.BackColor = &H80FF80
            End If
            
            If (IsDate(CStr(AsOfVal.Value)) = False) And DateOpt.Value = True Then
                Valid = False
                AsOfVal.BackColor = &H8080FF
            Else
                AsOfVal.BackColor = &H80FF80
            End If
            
            If Valid Then
                State = 1
                Run.Caption = "CONFIRM AND RUN"
            Else
                State = -1
                Run.Caption = "RETURN"
            End If
            
        Case -1
            
            AsOfVal.Locked = False
            LBalVal.Locked = False
            UBalVal.Locked = False
            TabLoc.Locked = False
            DateOpt.Locked = False
            BalOpt.Locked = False
            AsOfVal.BackColor = &H80000005
            LBalVal.BackColor = &H80000005
            UBalVal.BackColor = &H80000005
            TabLoc.BackColor = &H80000005
            
            For Each Entry In Cols
            
                Entry.BackColor = &H80000005
                Entry.Locked = False
                
            Next Entry
            
            State = 0
            Run.Caption = "VALIDATE"
            
        Case 1
        
            If BalOpt.Value = True Then
                Task = 1
            ElseIf DateOpt.Value = True Then
                Task = 0
            Else
                Task = -1
            End If
        
            If DateOpt.Value = True Then
                AsOfDate = CDate(CStr(AsOfVal.Value))
            Else
                AsOfDate = Date
            End If
            
            If BalOpt.Value = True Then
                AccL = CDbl(LBalVal.Value)
                AccH = CDbl(UBalVal.Value)
            Else
                AccL = 0#
                AccH = 0#
            End If
            
            RPTAB = CInt(TabLoc.Value)
            RPDIVJ = ColumnConverter(RPDIVJV.Value)
            RPDOC = ColumnConverter(RPDOCV.Value)
            RPAG = ColumnConverter(RPAGV.Value)
            RPDCT = ColumnConverter(RPDCTV.Value)
            RPDCTM = ColumnConverter(RPDCTMV.Value)
            RPGLBA = ColumnConverter(RPGLBAV.Value)
            RPDGJ = ColumnConverter(RPDGJV.Value)
            
            Unload Me
    End Select
    
End Sub


Private Sub UserForm_Activate()
    
    F_PaymentTool.Height = 146
    ResetParams
    
End Sub

Private Sub UserForm_Initialize()

    F_PaymentTool.Height = 146
    ResetParams
    
End Sub

Private Sub ResetParams()

    Fields = Array(AsOfVal, LBalVal, UBalVal, TabLoc, RPDIVJV, RPDOCV, RPAGV, RPDCTV, RPDCTMV, RPGLBAV, RPDGJV)

    State = 0

    DateOpt.Value = False
    BalOpt.Value = False
    DateOpt.Locked = False
    BalOpt.Locked = False
    
    For Each Entry In Fields
    
        Entry.Value = Null
        Entry.BackColor = &H80000005
        Entry.Locked = False
    
    Next Entry

    Run.Caption = "VALIDATE"

End Sub

Private Function BalRangeValidate(LBal As String, UBal As String) As Boolean

    If (Not IsNumeric(LBal)) Or (Not IsNumeric(UBal)) Then
        BalRangeValidate = False
        Exit Function
    End If
    
    If CDbl(LBal) > CDbl(UBal) Then
        BalRangeValidate = False
        Exit Function
    End If
    
    BalRangeValidate = True

End Function

Private Function ColumnValidate(Col As String) As Boolean
    
    If IsNumeric(Col) Then
        If (Col > 0) And (Col < 16384) Then
            ColumnValidate = True
            Exit Function
        End If
    End If
    
    If Len(Col) > 0 And Len(Col) < 4 Then
        
        ColumnValidate = IsLetter(Col)
        Exit Function
        
    End If
    
    ColumnValidate = False
    
End Function

Private Function ColumnConverter(Col As String) As Integer

    If IsNumeric(Col) Then
        ColumnConverter = CInt(Col)
        Exit Function
    End If
    
    ColumnConverter = Columns(CStr(Col)).Column
    
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Abort = True
    
    If CloseMode = 1 Then
        Abort = False
    End If
    
    ResetParams
    
End Sub
