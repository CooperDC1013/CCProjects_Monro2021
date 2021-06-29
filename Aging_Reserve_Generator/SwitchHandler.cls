VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwitchHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Tgl As MSForms.ToggleButton
Attribute Tgl.VB_VarHelpID = -1

Private Sub Tgl_Click()

    N = Tgl.Tag

    Vals = Array("OFF", "ON")
    Tgl.Caption = Vals(-1 * Tgl.Value)
    
    If Tgl.Value = False Then
        Tgl.BackColor = &H8080FF
        
        For Each Ctrl In GroupSelections.Major.Controls
        
            If Ctrl.Tag = N And TypeName(Ctrl) = "TextBox" Then
                Ctrl.Value = Null
            End If
            
        Next Ctrl
        
    Else
        Tgl.BackColor = &H80FF80
    End If

End Sub
