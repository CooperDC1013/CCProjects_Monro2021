VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadDef 
   Caption         =   "Load Definition"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "LoadDef.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Res As Collection

Private Sub Detect_Click()

    For i = Application.Worksheets.Count To 1 Step -1
        If Sheets(i).Visible = False Then
            Sheets(i).Visible = True
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Sheets(i).Visible = False
        End If
    Next i
    
    Dim Definitions As Dictionary
    Set Definitions = New Dictionary
    
    Dim Params As Collection
    
    If IsNumeric(TabB.Value) Then
        If TabB.Value > 0 And TabB.Value <= Application.Worksheets.Count Then
            If Sheets(CInt(TabB.Value)).Visible = True Then
                
                Sheets(CInt(TabB.Value)).Select
                
                Found = False
                
                For i = 1 To 30
                    For ii = 1 To 100
                    
                        If Cells(ii, i).Value = "<<<" And Cells(ii, i).Offset(2, 0).Value = ">>>" Then
                            
                            Header = Split(Cells(ii, i).Offset(1, 0), "%%%")(0)
                            HeaderVals = Split(Header, "%%")
                            
                            Set Params = New Collection
                            Params.Add (HeaderVals(0))
                            Params.Add (HeaderVals(1))
                            Params.Add (HeaderVals(2))
                            Params.Add Cells(ii, i).Offset(1, 0).Address
                            
                            Definitions.Add Key:=Cells(ii, i).Offset(1, 0).Value, Item:=Params
                            Found = True
                        End If
                    
                    Next ii
                Next i
                
                Discovered.Clear
                
                For Each Entry In Definitions.Keys
                
                    Discovered.AddItem
                    L = Discovered.ListCount
                    Discovered.List(L - 1, 0) = Definitions(Entry)(1)
                    Discovered.List(L - 1, 1) = Definitions(Entry)(2)
                    Discovered.List(L - 1, 2) = Definitions(Entry)(3)
                    Discovered.List(L - 1, 3) = Definitions(Entry)(4)
                
                Next Entry
                
            End If
        End If
    End If
    
End Sub

Private Sub Load_Click()
    
    If Discovered.ListIndex <> -1 And Discovered.ListCount > 0 Then
    
        Dim DefLoc As Range
        Set DefLoc = Range(CStr(Discovered.List(Discovered.ListIndex, 3)))
        
        Dim DataStr As String
        DataStr = CStr(DefLoc.Value)
        
        Dim Parsed As Collection
        Set Parsed = ParseDef(DataStr, 0)
        
        'Parsed is full data set. Now need to input into GroupSelections
        
        If TypeName(Parsed(2)) = "String" Then
            Set Res = New Collection
        Else
            Set Res = Parsed(2)
        End If
       
        Unload Me
        
    End If
        
End Sub

Private Function ParseDef(Inp As String, Level As Integer) As Collection

    Dim DataInp As String
    DataInp = Inp
    
    Dim Delims()
    Delims = Array("%%%", "%%", "@@", "&&", "^^^^^") 'Last delimiter is a cap that no string can be split by, ending the recursion.
    
    Dim Parsed() As String
    Parsed = Split(DataInp, Delims(Level))
    
    Dim Outer As Collection
    Set Outer = New Collection
    
    If UBound(Parsed) > 0 Then 'Initial string can be split
    
        Dim Inner As Collection
        
        For i = 0 To UBound(Parsed) 'For each part in split
        
            Set Inner = New Collection
        
            If UBound(Split(Parsed(i), Delims(Level + 1))) > 0 Then 'If part can be split on next delimiter then
                
                Set Inner = ParseDef(Parsed(i), (Level + 1)) 'Recursively run this function on part
                Outer.Add Inner
                
            Else 'Part cannot be split, just add to outer collection
                Outer.Add Parsed(i)
            End If
        
        Next i
        
    Else
        Outer.Add Parsed(0)
    End If
    
    Set ParseDef = Outer

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 1 Then
        Set GroupSelections.InputDef = Res
    Else
        Set Res = New Collection
        Set GroupSelections.InputDef = Res
        GroupSelections.LoadButton.BackColor = &H8000000F
        GroupSelections.LoadButton.Caption = "LOAD DEFINITION"
    End If
    
End Sub
