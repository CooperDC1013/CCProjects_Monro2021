VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveDef 
   Caption         =   "Save Definition"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   OleObjectBlob   =   "SaveDef.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    NameB.Value = Null
    DescB.Value = Null
    GroupSelections.SaveButton.BackColor = &H8000000F
    GroupSelections.SaveButton.Caption = "SAVE DEFINITION"
    SaveDef.Hide
End Sub

Private Sub Save_Click()

    GroupSelections.SaveButton.BackColor = &H80FF80
    GroupSelections.SaveButton.Caption = "SAVED"

    Dim Header As Collection
    Set Header = New Collection
        
    vName = NameB.Value
    vDesc = DescB.Value
    vDate = Date
    
    If vName = "" Then
        vName = "UNTITLED"
    End If
    
    If vDesc = "" Then
        vDesc = "N/A"
    End If
    
    Header.Add vName
    Header.Add vDesc
    Header.Add vDate
    
    Dim Result As String
    
    Dim SaveDefinition As Collection
    Set SaveDefinition = New Collection
    
    Dim GroupParams As Collection
    Set GroupParams = New Collection
    
    Dim NewGroup As Collection
    Dim NewCodes As Collection
    Dim NewCond As Boolean
    Dim NewDays As Long
    Dim NewAmt As Long
    Dim NewSwitch As Boolean
    
    For i = 0 To GroupSelections.NGroups
    
        Set NewGroup = New Collection
        Set NewCodes = New Collection
        
        Blank = False
        
        For Each Ctrl In GroupSelections.Major.Controls
        
            If IsNumeric(Ctrl.Tag) Then
            
                If CInt(Ctrl.Tag) = i Then
                
                    If TypeName(Ctrl) = "ListBox" Then
                    
                        If Ctrl.ListCount = 0 Then
                            Blank = True
                        End If
                        
                        For ii = 0 To Ctrl.ListCount - 1
                            NewCodes.Add Ctrl.List(ii)
                        Next ii
                        
                    ElseIf TypeName(Ctrl) = "ToggleButton" Then
                    
                        NewCond = Ctrl.Value
                    
                    ElseIf TypeName(Ctrl) = "OptionButton" Then
                        
                        If Ctrl.Caption = "AND" Then
                            NewSwitch = Ctrl.Value
                        End If
                        
                    ElseIf TypeName(Ctrl) = "TextBox" And (Left(Ctrl.Name, 3) = "Day") Then
                        
                        If IsNumeric(Ctrl.Value) Then
                            NewDays = CLng(Ctrl.Value)
                        Else
                            NewDays = 0
                        End If
                        
                    ElseIf TypeName(Ctrl) = "TextBox" And (Left(Ctrl.Name, 3) = "Amt") Then
                    
                        If IsNumeric(Ctrl.Value) Then
                            NewAmt = CLng(Ctrl.Value)
                        Else
                            NewAmt = 0
                        End If
                        
                    End If
                    
                End If
                
            End If
                
        Next Ctrl
            
        If Blank = False Then
            NewGroup.Add NewCodes
            NewGroup.Add NewCond
            NewGroup.Add NewDays
            NewGroup.Add NewAmt
            NewGroup.Add NewSwitch
        
            GroupParams.Add NewGroup
        End If
        
    Next i
    
    SaveDefinition.Add Header
    SaveDefinition.Add GroupParams
    
    Result = ConvertCollToStr(SaveDefinition, 0)
    
    Sheets.Add After:=Sheets(Application.Worksheets.Count)
    R = Application.Worksheets.Count
    
    Sheets(R).Name = vName
    
    Sheets(R).Select
    Range("A1").Value = "<<<"
    Range("A1").Offset(1, 0).Value = Result
    Range("A1").Offset(2, 0).Value = ">>>"
    
    With Range(Range("A1").Offset(4, 1), Range("A1").Offset(7, 2))
    
        .Cells(1).Value = "NAME"
        .Cells(3).Value = "DESCRIPTION"
        .Cells(5).Value = "DATE"
        
        .Cells(2).Value = vName
        .Cells(4).Value = vDesc
        .Cells(6).Value = vDate
    
    End With
    
    Dim NewStart As Range
    Set NewStart = Range("A1").Offset(10, 1)
    
    NGroup = 0
    
    For Each Group In GroupParams
    
        CodeCount = Group(1).Count
        
        With Range(NewStart, NewStart.Offset(5 + CodeCount, 1))
        
            .Cells(1).Value = "GROUP"
            .Cells(3).Value = "CONDITIONS"
            .Cells(5).Value = "DAYS"
            .Cells(7).Value = "AMT"
            .Cells(9).Value = "AND/OR"
            .Cells(11).Value = "CODES"
            
            .Cells(2).Value = NGroup
            NGroup = NGroup + 1
            
            .Cells(4).Value = Group(2)
            .Cells(6).Value = Group(3)
            .Cells(8).Value = Group(4)
            .Cells(10).Value = Group(5)
            
            For i = 1 To Group(1).Count
                .Cells(14 + ((i - 1) * 2)) = Group(1)(i)
                .Cells(14 + ((i - 1) * 2)).HorizontalAlignment = xlLeft
            Next i
            
            For i = 1 To 12
            
                If i Mod 2 = 0 Then
                    .Cells(i).HorizontalAlignment = xlCenter
                Else
                    .Cells(i).HorizontalAlignment = xlLeft
                End If
            
            Next i
        
        End With
        
        Set NewStart = NewStart.Offset(0, 3)
        
    Next Group
    
    NameB.Value = Null
    DescB.Value = Null
    
    Unload Me
    
End Sub

Private Function ConvertCollToStr(ByVal rData As Collection, Level As Integer) As String

    Delimiters = Array("%%%", "%%", "@@", "&&")
    
    Dim Data As Collection
    Set Data = rData
    
    Dim Res As String
    Res = ""
    
    Dim NewElement As String
    
    For Each Element In Data
    
        If TypeName(Element) = "Collection" Then
            NewElement = ConvertCollToStr((Element), (Level + 1))
            Res = Res & NewElement & Delimiters(Level)
        Else
            Res = Res & Element & Delimiters(Level)
        End If
    
    Next Element
    
    Res = Left(Res, Len(Res) - Len(Delimiters(Level)))
    
    ConvertCollToStr = Res
        
End Function


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1
        Cancel_Click
    End If
End Sub
