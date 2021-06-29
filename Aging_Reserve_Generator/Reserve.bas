Attribute VB_Name = "Reserve"
Public Abort As Boolean

Public Aging As Integer
Public CType As Integer
Public Address As Integer
Public Alpha As Integer
Public COpen As Integer
Public Curr As Integer
Public C30 As Integer
Public C60 As Integer
Public C90 As Integer
Public C120 As Integer
Public C150 As Integer

Public DataSet As Collection
Public AllGroups As Dictionary

Public Sub Reserve()

    For i = Application.Worksheets.Count To 1 Step -1
    
        If Sheets(i).Visible = False Then
        
            Sheets(i).Visible = True
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Sheets(Application.Worksheets.Count).Visible = False
            
        End If
        
    Next i
    
    Application.StatusBar = "COLUMN SELECTIONS"
    
    ColumnSelections.Show vbModal
    
    If Abort = True Then
        Exit Sub
    End If
    
    Cols = Array(CType, Address, Alpha, Curr, C30, C60, C90, C120, C150)
    MaxCol = 0
    MinCol = 1000000
    
    For i = 0 To UBound(Cols)
        
        If Cols(i) > MaxCol Then
            MaxCol = Cols(i)
        End If
        
        If Cols(i) < MinCol Then
            MinCol = Cols(i)
        End If
    
    Next i
   
    Application.StatusBar = "IDENTIFYING ALL CODES"
    DoEvents
    
    Dim TotalRows As Double
    
    Cells(Rows.Count, CType).End(xlUp).Select
    Range(Selection, Cells(2, CType).Address).Select
    
    TotalRows = Selection.Rows.Count
    
    Set DataSet = New Collection
    
    For Each Cell In Selection
    
        Exists = False
    
        For i = 1 To DataSet.Count
        
            If DataSet(i) = Cell.Value Then
                Exists = True
                Exit For
            End If
            
        Next i
        
        If Exists = False Then
            DataSet.Add Item:=Cell.Value
        End If
    
    Next Cell
    
    Application.StatusBar = "GROUP SELECTIONS"
    
    GroupSelections.Show vbModal
    
    Application.StatusBar = "CREATING GROUP AGING DETAIL"
    DoEvents
    
    Dim GroupDetails As Dictionary
    Dim Isolate As Dictionary
    Dim Codes As Collection
    Dim Prev As String
    Dim PrevF As String
    
    Dim ColOff As Integer
    Dim LastRow As Double
    Dim Copied As Double
    Dim Manual As Double
    Copied = 0
    Manual = 0
    
    Set GroupDetails = New Dictionary
    Set Isolate = New Dictionary
    
    For Each Key In AllGroups.Keys
    
        Application.Sheets.Add After:=Sheets(Application.Worksheets.Count)
        GroupDetails.Add Key:=Key, Item:=Sheets(Application.Worksheets.Count)
        Application.Sheets.Add After:=Sheets(Application.Worksheets.Count)
        Isolate.Add Key:=Key, Item:=Sheets(Application.Worksheets.Count)
        
        GroupDetails(Key).Name = "GROUP " & Key
        Isolate(Key).Name = "G" & Key & " ACCOUNTS"
        
        Set Codes = AllGroups(Key)(1)
        LastRow = 2
        Prev = ""
        
        Sheets(Aging).Select
        
        Rows(1).Copy GroupDetails(Key).Rows(1)
        
        Cells(Rows.Count, MinCol).End(xlUp).Select
        Range(Selection, Cells(2, MaxCol).Address).Select
        
        ColOff = Selection.Columns(1).Column - Columns(1).Column
        
        For Each Row In Selection.Rows
        
            Move = False
            Continue = True
            
            If Row.Cells(CType - ColOff).Value = Prev Then
                Move = True
            End If
            
            If Row.Cells(CType - ColOff).Value = PrevF Then
                Continue = False
            End If
            
            If Move = False And Continue = True Then
                Manual = Manual + 1
                NotFound = True
                
                For i = 1 To Codes.Count
                    
                    If Row.Cells(CType - ColOff).Value = Codes(i) Then
                        Move = True
                        NotFound = False
                        Prev = Row.Cells(CType - ColOff).Value
                    End If
                    
                Next i
                
                If NotFound = True Then
                    PrevF = Row.Cells(CType - ColOff).Value
                End If
                
            End If
            
            If Move = True Then
            
                'Row.EntireRow.Copy GroupDetails(Key).Rows(LastRow)
                GroupDetails(Key).Range(Cells(LastRow, 1).Address, Cells(LastRow, MaxCol - MinCol + 1).Address).Value = Row.Value
                LastRow = LastRow + 1
                Copied = Copied + 1
                
            End If
            
            If Copied <> 0 And Copied Mod 1000 = 0 Then
                DoEvents
                Application.StatusBar = "TOTAL ACCOUNTS ASSIGNED : " & Copied & " / " & TotalRows & " " & "MANUAL LOOKUPS : " & Manual & " CURRENT GROUP : " & Key & " ROW : " & Row.Row
                DoEvents
            End If
        
        Next Row
    
    Next Key
    
    Dim TestConditions As Boolean
    Dim TestDays As Integer
    Dim DayIndex As Integer
    Dim TestAmt As Double
    Dim TestOpt As Boolean
    
    Tot0 = Array(COpen)
    Tot150 = Array(C150)
    Tot120 = Array(C120, C150)
    Tot90 = Array(C90, C120, C150)
    Tot60 = Array(C60, C90, C120, C150)
    Tot30 = Array(C30, C60, C90, C120, C150)
    
    DayOpt = Array(Tot0, Tot30, Tot60, Tot90, Tot120, Tot150)
    
    ColOff = MinCol - 1
    
    Dim GroupAccts As Dictionary
    Dim PasterG As Double
    Dim Qualify As Boolean
    
    For Each Key In AllGroups.Keys
        
        Application.StatusBar = "ISOLATING GROUP " & Key & " ACCOUNTS"
        
        TestConditions = AllGroups(Key)(2)
        TestDays = CInt(AllGroups(Key)(3))
        TestAmt = CDbl(AllGroups(Key)(4))
        TestOpt = AllGroups(Key)(5)
        
        DayIndex = TestDays / 30
        
        Set GroupAccts = New Dictionary
        
        If TestConditions = True Then
        
            GroupDetails(Key).Select
            
            Cells(Rows.Count, MinCol - ColOff).End(xlUp).Select
            Range(Selection, Cells(2, MaxCol - ColOff).Address).Select
            
            If Selection.Rows.Count = 2 And Selection.Rows(1).Row = 1 Then
                AbortFor = True
            End If
            
            For Each Row In Selection.Rows
            
                If Not GroupAccts.Exists(Row.Cells(Address - ColOff).Value) Then
                    GroupAccts.Add Key:=Row.Cells(Address - ColOff).Value, Item:=Array(Row.Cells(CType - ColOff).Value, Row.Cells(Alpha - ColOff).Value, 0#, 0#)
                End If
                
                CBal = 0
                
                For i = 0 To UBound(DayOpt(DayIndex))
                    CBal = CBal + Row.Cells(DayOpt(DayIndex)(i) - ColOff).Value
                Next i
                
                uArray = GroupAccts(Row.Cells(Address - ColOff).Value)
                uArray(2) = uArray(2) + Row.Cells(COpen - ColOff).Value
                uArray(3) = uArray(3) + CBal
                GroupAccts.Item(Row.Cells(Address - ColOff).Value) = uArray
            
            Next Row
            
            Isolate(Key).Select
            PasterG = 2
            
            With Range("A1:I1")
            
                .Cells(1).Value = "ACCT TYPE"
                .Cells(2).Value = "ACCT #"
                .Cells(3).Value = "ACCT NAME"
                .Cells(4).Value = "TOTAL BAL"
                .Cells(5).Value = "CONDITIONAL BAL"
                .Cells(7).Value = "OVER:=" & TestDays
                .Cells(9).Value = "AMT:=" & TestAmt
                
                If TestOpt = True Then
                    .Cells(8).Value = "OPT:=AND"
                Else
                    .Cells(8).Value = "OPT:=OR"
                End If
                
                .Cells.HorizontalAlignment = xlLeft
                .Cells.Font.Bold = True
                
            End With
            
            For Each Acct In GroupAccts.Keys
            
                Qualify = False
                
                If TestOpt = True Then
                    
                    If GroupAccts(Acct)(3) >= TestAmt Then
                        Qualify = True
                    End If
                    
                Else
                
                    If (GroupAccts(Acct)(2) >= TestAmt) Or (GroupAccts(Acct)(3) > 0) Then
                        Qualify = True
                    End If
                    
                End If
                
                If Qualify = True Then
                
                    Cells(PasterG, 1).Value = GroupAccts(Acct)(0)
                    Cells(PasterG, 2).Value = Acct
                    Cells(PasterG, 3).Value = GroupAccts(Acct)(1)
                    Cells(PasterG, 4).Value = GroupAccts(Acct)(2)
                    Cells(PasterG, 5).Value = GroupAccts(Acct)(3)
                    
                    Rows(PasterG).HorizontalAlignment = xlLeft
                    
                    PasterG = PasterG + 1
                
                End If
            
            Next Acct
            
            Columns(4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Columns(5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            Range(Columns(1), Columns(10)).AutoFit
            Range("A1").Select
            
        Else
        
            Application.DisplayAlerts = False
            Isolate(Key).Delete
            Application.DisplayAlerts = True
        
        End If
    
    Next Key
    
    For i = 1 To Application.Worksheets.Count
    
        Sheets(i).Select
        Range("A1").Select
        
    Next i
    
    Sheets(Aging).Select
    Application.StatusBar = False
    
    Finish = MsgBox("Execution Complete", vbOKOnly, "RESERVE")

End Sub
