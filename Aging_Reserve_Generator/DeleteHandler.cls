VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Cmd As MSForms.CommandButton
Attribute Cmd.VB_VarHelpID = -1

Private Sub Cmd_Click()

    Dim DeleteLast As Collection
    Set DeleteLast = New Collection

    N = Cmd.Tag
    
    GroupSelections.HOff = Not GroupSelections.HOff
    GroupSelections.NGroups = GroupSelections.NGroups - 1
    GroupSelections.TotalGroups.Caption = "TOTAL GROUPS" & vbCr & vbCr & GroupSelections.NGroups
    
     If GroupSelections.State = 1 Then
        GroupSelections.State = 0
        GroupSelections.Ok.Caption = "OK"
        GroupSelections.Ok.BackColor = &H8000000F
    End If
    
    If NGroups Mod 2 = 1 Then
        Major.Height = Major.Height - 252
        Major.ScrollHeight = Major.Height + (252 * (NGroups \ 2))
    End If
    
    For Each Ctrl In GroupSelections.Major.Controls
    
        If Ctrl.Tag = N And TypeName(Ctrl) = "Frame" Then
            
            DeleteLast.Add Ctrl
            
        ElseIf Ctrl.Tag = N And TypeName(Ctrl) <> "Frame" Then
        
            If TypeName(Ctrl) = "ListBox" Then
                
                For i = Ctrl.ListCount - 1 To 0 Step -1
                    GroupSelections.AllTypes.AddItem
                    GroupSelections.AllTypes.List(GroupSelections.AllTypes.ListCount - 1) = Ctrl.List(i)
                    Ctrl.RemoveItem (i)
                Next i
                
                Alpha GroupSelections.AllTypes
                
            End If
        
            GroupSelections.Major.Controls.Remove Ctrl.Name
        
        ElseIf Ctrl.Tag > N Then
            
            If TypeName(Ctrl) = "Label" Then
                
                If Left(Ctrl.Caption, 5) = "Group" Then
                    Ctrl.Caption = Replace(Ctrl.Caption, Mid(Ctrl.Caption, 8, 1), CStr(CInt(Mid(Ctrl.Caption, 8, 1)) - 1))
                End If
                
            End If
            
            Ctrl.Tag = Ctrl.Tag - 1
            
            If TypeName(Ctrl.Parent.Parent) = "GroupSelections" Then
            
                If Ctrl.Left > 340 Then
                    Ctrl.Left = Ctrl.Left - 340
                Else
                    Ctrl.Left = Ctrl.Left + 340
                    Ctrl.Top = Ctrl.Top - 252
                End If
            
            End If
            
        End If
    
    Next Ctrl
    
    For i = 1 To DeleteLast.Count
        GroupSelections.Major.Controls.Remove DeleteLast(i).Name
    Next i

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

Private Function Flr(Inp As Double, Optional Sig As Integer = 1) As Integer

    Dim DInp As Integer
    DInp = Fix(Inp)

    Do While DInp Mod Sig <> 0
        DInp = DInp - 1
    Loop
    
    Flr = DInp

End Function

