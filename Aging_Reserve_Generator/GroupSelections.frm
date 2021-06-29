VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupSelections 
   Caption         =   "Select Groups"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18630
   OleObjectBlob   =   "GroupSelections.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HOff As Boolean
Public NGroups As Integer
Public State As Integer
Private Handlers As Collection
Private InEvents As InHandler
Private OutEvents As OutHandler
Private DeleteEvents As DeleteHandler
Private SwitchEvents As SwitchHandler
Private AmtEvents As AmtHandler
Private DaysEvents As DaysHandler

Public InputDef As Collection

Private Sub AddGroup_Click()

    NGroups = NGroups + 1 'Use this new number to assign to tags of new controls to edit easier later
    TotalGroups.Caption = "TOTAL GROUPS" & vbCr & vbCr & NGroups
    SGroups = Str(NGroups)
    HOff = Not HOff 'Reverse sign to determine which side of frame to place on
    VOff = VOff + 1 'This int multiplied by 240 gives the vertical offset in the frame
    
    If State = 1 Then
        State = 0
        Ok.Caption = "OK"
        Ok.BackColor = &H8000000F
    End If
    
    L1 = "Isolate accounts when the following conditions are met :"
    L2 = "Has balance present over"
    L3 = "days"
    L4 = "Where such balance is >="
    L5 = "$"
    
    Dim AddList As Control
    Dim AddIn As Control
    Dim AddOut As Control
    Dim AddDelete As Control
    Dim AddFrame As Control
    Dim AddTitle As Control
    Dim AddSwitch As Control
    Dim AddFL1 As Control
    Dim AddFL2 As Control
    Dim AddFL3 As Control
    Dim AddFL4 As Control
    Dim AddFL5 As Control
    Dim AddDays As Control
    Dim AddAmt As Control
    Dim AddAnd As Control
    Dim AddOr As Control
    
    Set InEvents = New InHandler
    Set OutEvents = New OutHandler
    Set DeleteEvents = New DeleteHandler
    Set SwitchEvents = New SwitchHandler
    Set AmtEvents = New AmtHandler
    Set DaysEvents = New DaysHandler
    
    Handlers.Add InEvents
    Handlers.Add OutEvents
    Handlers.Add DeleteEvents
    Handlers.Add SwitchEvents
    Handlers.Add DaysEvents
    Handlers.Add AmtEvents
    
    Set AddList = Major.Controls.Add("Forms.Listbox.1", "Type" & SGroups)
    With AddList
        .Height = 184
        .Width = 70
        .Left = 78
        .Top = 54
        .Tag = NGroups
        .Font.Name = "Tahoma"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    Set AddIn = Major.Controls.Add("Forms.Commandbutton.1", "In" & SGroups)
    Set InEvents.Cmd = AddIn
    With AddIn
        .Height = 24
        .Width = 60
        .Left = 12
        .Top = 60
        .Tag = NGroups
        .Caption = "IN >"
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddOut = Major.Controls.Add("Forms.Commandbutton.1", "Out" & SGroups)
    Set OutEvents.Cmd = AddOut
    With AddOut
        .Height = 24
        .Width = 60
        .Left = 12
        .Top = 90
        .Tag = NGroups
        .Caption = "OUT <"
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddDelete = Major.Controls.Add("Forms.Commandbutton.1", "Delete" & SGroups)
    Set DeleteEvents.Cmd = AddDelete
    With AddDelete
        .Height = 54
        .Width = 60
        .Left = 12
        .Top = 120
        .Tag = NGroups
        .Caption = "DELETE GROUP"
        .WordWrap = True
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFrame = Major.Controls.Add("Forms.Frame.1", "Frame" & SGroups)
    With AddFrame
        .Height = 156
        .Width = 162
        .Left = 156
        .Top = 48
        .Tag = NGroups
        .Caption = "Account Conditions"
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddTitle = Major.Controls.Add("Forms.Label.1", "Title" & SGroups)
    With AddTitle
        .Height = 18
        .Width = 120
        .Left = 18
        .Top = 18
        .Tag = NGroups
        .Caption = "Group " & SGroups & " :"
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddSwitch = AddFrame.Controls.Add("Forms.ToggleButton.1", "Switch" & SGroups)
    Set SwitchEvents.Tgl = AddSwitch
    With AddSwitch
        .Height = 25
        .Width = 36
        .Left = 114
        .Top = 12
        .Tag = NGroups
        .Caption = "ON"
        .Value = True
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFL1 = AddFrame.Controls.Add("Forms.Label.1", "L1" & SGroups)
    With AddFL1
        .Height = 36
        .Width = 102
        .Left = 6
        .Top = 12
        .Tag = NGroups
        .Caption = L1
        .Visible = True
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFL2 = AddFrame.Controls.Add("Forms.Label.1", "L2" & SGroups)
    With AddFL2
        .Height = 24
        .Width = 60
        .Left = 6
        .Top = 60
        .Tag = NGroups
        .Caption = L2
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFL3 = AddFrame.Controls.Add("Forms.Label.1", "L3" & SGroups)
    With AddFL3
        .Height = 12
        .Width = 24
        .Left = 114
        .Top = 72
        .Tag = NGroups
        .Caption = L3
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFL4 = AddFrame.Controls.Add("Forms.Label.1", "L4" & SGroups)
    With AddFL4
        .Height = 24
        .Width = 60
        .Left = 6
        .Top = 96
        .Tag = NGroups
        .Caption = L4
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddFL5 = AddFrame.Controls.Add("Forms.Label.1", "L5" & SGroups)
    With AddFL5
        .Height = 18
        .Width = 12
        .Left = 78
        .Top = 105
        .Tag = NGroups
        .Caption = L5
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddDays = AddFrame.Controls.Add("Forms.TextBox.1", "Days" & SGroups)
    Set DaysEvents.Txt = AddDays
    With AddDays
        .Height = 20
        .Width = 36
        .Left = 72
        .Top = 66
        .Tag = NGroups
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddAmt = AddFrame.Controls.Add("Forms.TextBox.1", "Amt" & SGroups)
    Set AmtEvents.Txt = AddAmt
    With AddAmt
        .Height = 20
        .Width = 66
        .Left = 90
        .Top = 102
        .Tag = NGroups
        .Font.Name = "Tahoma"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddAnd = AddFrame.Controls.Add("Forms.OptionButton.1", "And" & SGroups)
    With AddAnd
        .Height = 16
        .Width = 45
        .Left = 6
        .Top = 126
        .Tag = NGroups
        .GroupName = NGroups
        .Value = True
        .Caption = "AND"
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    Set AddOr = AddFrame.Controls.Add("Forms.OptionButton.1", "Or" & SGroups)
    With AddOr
        .Height = 16
        .Width = 45
        .Left = 56
        .Top = 126
        .Tag = NGroups
        .GroupName = NGroups
        .Value = False
        .Caption = "OR"
        .Font.Name = "Tahoma"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Weight = 200
    End With
    
    For Each Ctrl In Major.Controls
        
        If Ctrl.Tag = NGroups And Ctrl.Parent.Name = "Major" Then
        
            If HOff = True Then
                Ctrl.Left = Ctrl.Left + 340
            End If
            
            Ctrl.Top = Ctrl.Top + (252 * (NGroups \ 2))
            
        End If
        
    Next Ctrl
    
    If NGroups Mod 2 = 0 Then
        Major.Height = Major.Height + 252
        Major.ScrollHeight = Major.Height + (252 * (NGroups \ 2))
    End If
    
End Sub

Private Sub Clear_Click()
    
    NGroups = 0
    TotalGroups.Caption = "TOTAL GROUPS" & vbCr & vbCr & NGroups
    HOff = False
    
    SaveButton.BackColor = &H8000000F
    LoadButton.BackColor = &H8000000F
    SaveButton.Caption = "SAVE DEFINITION"
    LoadButton.Caption = "LOAD DEFINITION"
    
    If State = 1 Then
        State = 0
        Ok.BackColor = &H8000000F
        Ok.Caption = "OK"
    End If
    
    For i = Handlers.Count To 1 Step -1
        Handlers.Remove (i)
    Next i
    
    Dim DeleteLast As Collection
    Set DeleteLast = New Collection
    
    For Each Ctrl In Major.Controls
    
        If IsNumeric(Ctrl.Tag) And (TypeName(Ctrl) <> "Frame") And (Ctrl.Tag <> "0") Then
            Major.Controls.Remove Ctrl.Name
        ElseIf IsNumeric(Ctrl.Tag) And TypeName(Ctrl) = "Frame" And (Ctrl.Tag <> "0") Then
            DeleteLast.Add Ctrl
        End If
        
    Next Ctrl
    
    For i = 1 To DeleteLast.Count
        Major.Controls.Remove DeleteLast(i).Name
    Next i
    
    CondSwitch.Value = False
    CondSwitch.Caption = "OFF"
    CondSwitch.BackColor = &H8080FF
    Amt0.Value = Null
    Days0.Value = Null
    
    Major.Height = 252
    Major.ScrollHeight = 252
    
    AllTypes.Clear
    
    For i = 1 To Reserve.DataSet.Count
    
        AllTypes.AddItem
        AllTypes.List(AllTypes.ListCount - 1) = Reserve.DataSet(i)
        
    Next i
    
    Alpha AllTypes
    
End Sub

Private Sub CommandButton1_Click()
    Testy
End Sub

Private Sub Amt0_AfterUpdate()
    If IsNumeric(Amt0.Value) Then
    
        If Amt0.Value < 0 Then
            Amt0.Value = 0
        Else
            Amt0.Value = Flr(Amt0.Value, 50)
        End If
        
    Else
        Amt0.Value = 0
    End If
End Sub

Private Sub CondSwitch_Click()
    Vals = Array("OFF", "ON")
    CondSwitch.Caption = Vals(-1 * CondSwitch.Value)
    
    If CondSwitch.Value = False Then
        CondSwitch.BackColor = &H8080FF
        Amt0.Value = Null
        Days0.Value = Null
    Else
        CondSwitch.BackColor = &H80FF80
    End If
        
End Sub

Private Sub CondTypes_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CondTypes.Clear
    
    For i = 0 To AllTypes.ListCount - 1
        CondTypes.AddItem
        CondTypes.List(i) = AllTypes.List(i)
    Next i
    
End Sub

Private Sub LoadButton_Click()
        
    LoadDef.Show
    Debug.Print ("CONTROL RETURNED")
    
    If InputDef.Count <> 0 Then
    
        Clear_Click
        
        Dim Codes As Collection
    
        For i = 2 To InputDef.Count
        
            AddGroup_Click
            
            Set Codes = InputDef(i)(1)
            
            For Each Ctrl In Major.Controls
            
                If IsNumeric(Ctrl.Tag) Then
            
                    If CInt(Ctrl.Tag) = (i - 1) Then
                    
                        If TypeName(Ctrl) = "ListBox" Then
                        
                            For Each C In Codes
                            
                                For ii = 0 To (AllTypes.ListCount - 1)
                                
                                    If AllTypes.List(ii) = CStr(C) Then
                                        AllTypes.ListIndex = ii
                                        Ctrl.AddItem
                                        Ctrl.List(Ctrl.ListCount - 1) = AllTypes.List(ii)
                                        AllTypes.RemoveItem (ii)
                                        Exit For
                                    End If
                                    
                                Next ii
                            
                            Next C
                            
                        ElseIf TypeName(Ctrl) = "ToggleButton" Then
                        
                            Ctrl.Value = CBool(InputDef(i)(2))
                            
                        ElseIf TypeName(Ctrl) = "TextBox" Then
                        
                            If Left(Ctrl.Name, 3) = "Day" Then
                            
                                Ctrl.Value = CLng(InputDef(i)(3))
                            
                            ElseIf Left(Ctrl.Name, 3) = "Amt" Then
                                
                                Ctrl.Value = CLng(InputDef(i)(4))
                                
                            End If
                            
                        ElseIf TypeName(Ctrl) = "OptionButton" Then
                        
                            If Ctrl.Caption = "AND" Then
                            
                                Ctrl.Value = CBool(InputDef(i)(5))
                                
                            ElseIf Ctrl.Caption = "OR" Then
                            
                                Ctrl.Value = Not CBool(InputDef(i)(5))
                                
                            End If
                            
                        End If
                        
                    End If
                
                End If
            
            Next Ctrl
        
        Next i
        
        AllTypes.ListIndex = 0
        
        LoadButton.BackColor = &H80FF80
        LoadButton.Caption = "LOADED"
    
    End If
    
End Sub

Private Sub Ok_Click()
    
    If State = 0 Then
    
        Ok.BackColor = &H80FF80
        Ok.Caption = "RUN"
        State = 1
    
    ElseIf State = 1 Then
    
        CondTypes.Clear
    
        For i = 0 To AllTypes.ListCount - 1
            CondTypes.AddItem
            CondTypes.List(i) = AllTypes.List(i)
        Next i
    
        Dim SendGroups As Dictionary
        Set SendGroups = New Dictionary
        
        Dim GroupData As Collection
        Dim GroupCodes As Collection
        Dim GroupSwitch As Boolean
        Dim GroupDays As Integer
        Dim GroupAmt As Integer
        Dim GroupAnd As Boolean
        
        MaxTag = -1
        
        For Each Ctrl In Major.Controls
            If IsNumeric(Ctrl.Tag) Then
                If CInt(Ctrl.Tag) > MaxTag Then
                    MaxTag = CInt(Ctrl.Tag)
                End If
            End If
        Next Ctrl
        
        For i = 0 To MaxTag
        
            Blank = False
        
            For Each Ctrl In Major.Controls
            
                If IsNumeric(Ctrl.Tag) Then
                
                    If CInt(Ctrl.Tag) = i Then
                    
                        Select Case TypeName(Ctrl)
                            Case "ListBox"
                            
                                If Ctrl.ListCount = 0 Then
                                    Blank = True
                                End If
                            
                                Set GroupCodes = New Collection
                                
                                For ii = 0 To Ctrl.ListCount - 1
                                    GroupCodes.Add Ctrl.List(ii)
                                Next ii
                                
                            Case "ToggleButton"
                                GroupSwitch = Ctrl.Value
                            Case "OptionButton"
                            
                                If Ctrl.Caption = "AND" Then
                                    GroupAnd = Ctrl.Value
                                End If
                            
                            Case "TextBox"
                            
                                If Left(Ctrl.Name, 3) = "Day" Then
                                    If IsNumeric(Ctrl.Value) Then
                                        GroupDays = Ctrl.Value
                                    Else
                                        GroupDays = 0
                                    End If
                                ElseIf Left(Ctrl.Name, 3) = "Amt" Then
                                    If IsNumeric(Ctrl.Value) Then
                                        GroupAmt = Ctrl.Value
                                    Else
                                        GroupAmt = 0
                                    End If
                                End If
                        End Select
                        
                    End If
                    
                End If
                
            Next Ctrl
            
            If Blank = False Then
            
                Set GroupData = New Collection
                
                GroupData.Add GroupCodes
                GroupData.Add GroupSwitch
                GroupData.Add GroupDays
                GroupData.Add GroupAmt
                GroupData.Add GroupAnd
                
                SendGroups.Add Key:=i, Item:=GroupData
                
            End If
                        
        Next i
        
        Set Reserve.AllGroups = SendGroups
        
        Unload Me

    End If
    
End Sub

Private Sub PgDn_Click()
    If Major.ScrollTop + 252 > Major.ScrollHeight Then
        Major.ScrollTop = Major.ScrollHeight
    Else
        Major.ScrollTop = Major.ScrollTop + 252
    End If
    
    AllTypes.SetFocus
End Sub

Private Sub PgUp_Click()
    If Major.ScrollTop - 252 < 0 Then
        Major.ScrollTop = 0
    Else
        Major.ScrollTop = Major.ScrollTop - 252
    End If
    
    AllTypes.SetFocus
End Sub

Private Sub Days0_AfterUpdate()
    
    If IsNumeric(Days0.Value) Then
    
        If Days0.Value > 150 Then
            Days0.Value = 150
        ElseIf Days0.Value < 0 Then
            Days0.Value = 0
        Else
            Days0.Value = Flr(Days0.Value, 30)
        End If
        
    Else
        Days0.Value = 0
    End If
    
End Sub

Private Sub SaveButton_Click()
    
    CondTypes.Clear
    
        For i = 0 To AllTypes.ListCount - 1
            CondTypes.AddItem
            CondTypes.List(i) = AllTypes.List(i)
        Next i
    
    SaveDef.Show
    
End Sub

Private Sub UserForm_Initialize()
    Set Handlers = New Collection
    NGroups = 0
    
    State = 0
    
    Set InputDef = New Collection
    
    TotalGroups.Caption = "TOTAL GROUPS" & vbCr & vbCr & NGroups
    
    AllTypes.Clear
    
    For i = 1 To Reserve.DataSet.Count
    
        AllTypes.AddItem
        AllTypes.List(AllTypes.ListCount - 1) = Reserve.DataSet(i)
        
    Next i
        
    Alpha AllTypes
    
    AddGroup.SetFocus
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Reserve.Abort = False
    
    If CloseMode <> 1 Then
        Reserve.Abort = True
        Set Reserve.AllGroups = New Dictionary
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
