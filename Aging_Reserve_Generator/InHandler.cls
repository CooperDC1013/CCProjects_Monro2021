VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Cmd As MSForms.CommandButton
Attribute Cmd.VB_VarHelpID = -1

Private Sub Cmd_Click()

    Dim ListBox As Control
    N = Cmd.Tag
    
    If GroupSelections.State = 1 Then
        GroupSelections.State = 0
        GroupSelections.Ok.Caption = "OK"
        GroupSelections.Ok.BackColor = &H8000000F
    End If
    
    For Each Ctrl In GroupSelections.Major.Controls
    
        If Ctrl.Tag = N And TypeName(Ctrl) = "ListBox" Then
            Set ListBox = Ctrl
        End If
        
    Next Ctrl
    
    If GroupSelections.AllTypes.ListIndex <> -1 Then
    
        ListBox.AddItem
        ListBox.List(ListBox.ListCount - 1) = GroupSelections.AllTypes.List(GroupSelections.AllTypes.ListIndex)
        GroupSelections.AllTypes.RemoveItem (GroupSelections.AllTypes.ListIndex)
        
    End If
    
    Alpha ListBox
    
    GroupSelections.AllTypes.SetFocus
    
End Sub

Private Sub Alpha(L As Control)

    Dim Sorted As Collection
    Set Sorted = New Collection
    
    For i = 0 To (L.ListCount - 1)
        Sorted.Add L.List(i)
    Next i
    
    Done = False
    
    Do While Done = False
        
        Done = True
    
        For i = 1 To (Sorted.Count - 1)
        
            If UCase(Sorted(i)) > UCase(Sorted(i + 1)) Then
                Done = False
                First = Sorted(i + 1)
                Sec = Sorted(i)
                Sorted.Remove (i)
                Sorted.Add Item:=First, Before:=i
                Sorted.Remove (i + 1)
                Sorted.Add Item:=Sec, After:=i
            End If
            
        Next i
        
    Loop
    
    For i = 1 To Sorted.Count
    
        L.List(i - 1) = Sorted(i)
        
    Next i

End Sub
