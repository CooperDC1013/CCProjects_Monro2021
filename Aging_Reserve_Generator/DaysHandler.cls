VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaysHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1

Private Sub Txt_AfterUpdate()

    If IsNumeric(Txt.Value) Then
        Txt.Value = Flr(Txt.Value, 30)
    Else
        Txt.Value = 0
    End If

End Sub

Private Sub Txt_Change()

End Sub

Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsNumeric(Txt.Value) Then
    
        If Txt.Value > 150 Then
            Txt.Value = 150
        ElseIf Txt.Value < 0 Then
            Txt.Value = 0
        Else
            Txt.Value = Flr(Txt.Value, 30)
        End If
        
    Else
        Txt.Value = 0
    End If
End Sub

Private Function Flr(Inp As Double, Optional Sig As Long = 1) As Long

    Dim DInp As Long
    DInp = Fix(Inp)

    Do While DInp Mod Sig <> 0
        DInp = DInp - 1
    Loop
    
    Flr = DInp

End Function
