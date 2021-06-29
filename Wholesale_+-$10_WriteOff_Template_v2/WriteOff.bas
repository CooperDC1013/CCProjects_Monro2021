Attribute VB_Name = "WriteOff"
Public Spread As Boolean
Public DocTypeCol As Integer
Public AccountCol As Integer
Public DateCol As Integer
Public InvoiceCol As Integer
Public OpenCol As Integer
Public GrossCol As Integer
Public BUCol As Integer
Public BU3Col As Integer
Public BU5Col As Integer

Public DetailTab As Integer

Public GLD As Date

Public CycleVar As Integer

'W - Spread
'Q - Write Off

Private Sub ChangeCycle()

    If CycleVar = 6 Then
        CycleVar = 0
    Else
        CycleVar = CycleVar + 1
    End If

End Sub


Public Sub WriteOff10()

    For i = Application.Worksheets.Count To 1 Step -1
    
        If Sheets(i).Visible = False Then
            Sheets(i).Visible = True
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Sheets(Application.Worksheets.Count).Visible = False
        End If
        
    Next i
    
    AgingForm.Show vbModal

    Dim Minor As New Dictionary
    
    CycleVar = 0
    Cycle = Array(1431, 1432, 1433, 1434, 1537, 1538, 1166)
    Dim GL As Dictionary
    Set GL = New Dictionary
    
    GL.Add Key:=1431, Item:="00508366"
    GL.Add Key:=1432, Item:="00508412"
    GL.Add Key:=1433, Item:="00508458"
    GL.Add Key:=1434, Item:="00508504"
    GL.Add Key:=1537, Item:="00549803"
    GL.Add Key:=1538, Item:="00550062"
    GL.Add Key:=1539, Item:="00550321"
    GL.Add Key:=1166, Item:="00366840"

    If Spread = True Then

        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        W = Application.Worksheets.Count
        Sheets(W).Name = "$10 SPREAD"
        PasterW = 1
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        Q = Application.Worksheets.Count
        Sheets(Q).Name = "$10 WRITE OFFS"
        PasterQ = 1

        Sheets(DetailTab).Select

        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        Dim AcctBal As New Dictionary
        Set AcctBal = New Dictionary
        
        Application.StatusBar = "CALCULATING BALANCES"
        DoEvents
        
        For Each Row In Selection.Rows
        
            If Not AcctBal.Exists(Row.Cells(AccountCol).Value) Then
                AcctBal.Add Key:=CDbl(Row.Cells(AccountCol).Value), Item:=Array(0#, 0, True, False)
            End If
            
            Count = AcctBal(Row.Cells(AccountCol).Value)(1)
            Bal = AcctBal(Row.Cells(AccountCol).Value)(0)
            NewBal = Bal + Row.Cells(OpenCol).Value

            AcctBal(Row.Cells(AccountCol).Value) = Array(NewBal, Count + 1, True, False)
            
        Next Row
        
        For Each Row In Selection.Rows
        
            Count = AcctBal(Row.Cells(AccountCol).Value)(1)
            Bal = AcctBal(Row.Cells(AccountCol).Value)(0)
            SameSign = AcctBal(Row.Cells(AccountCol).Value)(2)
            SameBal = AcctBal(Row.Cells(AccountCol).Value)(3)
            
            If Bal <> 0 Then
            
                If ((Row.Cells(OpenCol).Value / Bal) < 0) And SameSign = True Then
                    SameSign = False
                End If
                
            End If
            
            AcctBal(Row.Cells(AccountCol).Value) = Array(Bal, Count, SameSign, SameBal)
            
        Next Row
        
        Dim AdjustSpread As Dictionary
        Set AdjustSpread = New Dictionary
        Dim Skip As Collection
        Set Skip = New Collection
        
        Application.StatusBar = "PRIMARY ASSIGNMENT"
        DoEvents
        
        For Each Row In Selection.Rows
        
            Bal = AcctBal(Row.Cells(AccountCol).Value)(0)
            Count = AcctBal(Row.Cells(AccountCol).Value)(1)
            SameSign = AcctBal(Row.Cells(AccountCol).Value)(2)
            SameBal = AcctBal(Row.Cells(AccountCol).Value)(3)
        
            If Bal < 10# And Bal > -10# Then
            
                If Round(Bal, 2) = 0 Then
                
                    Row.EntireRow.Copy Sheets(W).Range(PasterW & ":" & PasterW)
                    PasterW = PasterW + 1
                    
                ElseIf Count = 1 Then
                
                    Row.EntireRow.Copy Sheets(Q).Range(PasterQ & ":" & PasterQ)
                    PasterQ = PasterQ + 1
                    
                ElseIf SameSign = True Then
                
                    Row.EntireRow.Copy Sheets(Q).Range(PasterQ & ":" & PasterQ)
                    PasterQ = PasterQ + 1
                    
                ElseIf (Row.Cells(OpenCol).Value = Bal) And (SameBal = False) Then
                
                    Row.EntireRow.Copy Sheets(Q).Range(PasterQ & ":" & PasterQ)
                    PasterQ = PasterQ + 1
                    
                    AcctBal(Row.Cells(AccountCol).Value) = Array(Bal, Count, SameSign, True)
                    Skip.Add (Row.Cells(AccountCol).Value)
                    
                Else
                
                    Row.EntireRow.Copy Sheets(W).Range(PasterW & ":" & PasterW)
                    PasterW = PasterW + 1
                    
                    If Not AdjustSpread.Exists(Row.Cells(AccountCol).Value) Then
                        AdjustSpread.Add Key:=Row.Cells(AccountCol).Value, Item:=Array(Bal, True)
                    End If
                    
                End If
                
            End If
            
        Next Row
        
        Application.StatusBar = "SECONDARY ASSIGNMENT"
        
        Sheets(W).Select
        
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("1:1")).Select
        
        For Each Row In Selection.Rows
        
            SkipMe = False
            
            For i = 1 To Skip.Count
            
                If Row.Cells(AccountCol).Value = Skip(i) Then
                    SkipMe = True
                End If
                
            Next i
        
            If (AdjustSpread.Exists(Row.Cells(AccountCol).Value)) And (SkipMe = False) Then
            
                Bal = AdjustSpread(Row.Cells(AccountCol).Value)(0)
                Adjust = AdjustSpread(Row.Cells(AccountCol).Value)(1)
                RelSign = Row.Cells(OpenCol).Value / Bal
                
                If (RelSign > 0) And Adjust = True Then
                
                    PreviousAmt = Row.Cells(OpenCol).Value 'Track to revert to later
                    
                    Row.Cells(OpenCol).Value = Bal 'Adjust amount to match what will be written off
                    
                    Row.EntireRow.Copy Sheets(Q).Range(PasterQ & ":" & PasterQ) 'Copy to Write Off Sheet
                    PasterQ = PasterQ + 1
                    
                    Row.Cells(OpenCol).Value = PreviousAmt - Bal 'Adjust amount to match what actually will be spread Invoice - Bal
                    
                    Adjust = False 'Flip so we only write off one invoice per account
                    AdjustSpread(Row.Cells(AccountCol).Value) = Array(Bal, Adjust) 'Reinsert item into Adjustspread to reflect new Adjust boolean
                
                End If
            
            End If
        
        Next Row
        
        Application.StatusBar = "FORMATTING..."

        For i = 1 To 7
            Columns(1).Insert
        Next i

        Columns(InvoiceCol + 7).EntireColumn.Cut Sheets(W).Columns(1)
        Columns(DocTypeCol + 7).EntireColumn.Cut Sheets(W).Columns(2)
        Columns(BU3Col + 7).EntireColumn.Cut Sheets(W).Columns(3)
        Columns(OpenCol + 7).EntireColumn.Cut Sheets(W).Columns(4)
        Columns(BU5Col + 7).EntireColumn.Cut Sheets(W).Columns(6)
        Columns(AccountCol + 7).EntireColumn.Cut Sheets(W).Columns(7)

        Columns("D:D").EntireColumn.NumberFormat = "#,##0.00"
        Columns("H:BB").EntireColumn.Select
        Columns("H:BB").EntireColumn.Delete

        Columns("A:BB").EntireColumn.AutoFit

        For i = 1 To 15
            Columns(1).Insert
        Next i

        Cells(Rows.Count, 16).End(xlUp).EntireRow.Select
        Range(Selection, Range("1:1")).Select

        For Each Row In Selection.Rows

            Row.Cells(1).NumberFormat = "@"
            Row.Cells(1).Value = "00901"
            Row.Cells(2).Value = "I"
            Row.Cells(3).Value = "9"
            Row.Cells(4).Value = "1"
            Row.Cells(6).Value = Month(GLD)
            Row.Cells(7).Value = Day(GLD)
            Row.Cells(8).Value = Right(Year(GLD), 2)
            Row.Cells(9).Value = Left(Year(GLD), 2)
            Row.Cells(10).Value = Row.Cells(22).Value
            Row.Cells(11).Value = Month(GLD)
            Row.Cells(12).Value = Day(GLD)
            Row.Cells(13).Value = Right(Year(GLD), 2)
            Row.Cells(14).Value = Left(Year(GLD), 2)
            
            Row.Cells(15).NumberFormat = "@"
            Row.Cells(15).Value = "00040806"
            
            Row.Cells(20).Value = "0"

            For i = 1 To 4
                Row.Cells(i).Interior.ColorIndex = 22
            Next i

            Row.Cells(5).Interior.ColorIndex = 40

            For i = 6 To 15
                Row.Cells(i).Interior.ColorIndex = 39
            Next i

            For i = 16 To 22
                Row.Cells(i).Interior.ColorIndex = 37
            Next i

        Next Row

        Range(Columns(1), Columns(22)).ColumnWidth = 8.43

        For i = 1 To 5

            Range("A1").EntireRow.Insert

        Next i
        
        Sheets(W).Select
    
        Header = InputBox("Enter Tab # where AUTOCASH header is located.", "Insert Header?", DetailTab + 1)
    
        If (Header <> "") And (IsNumeric(Header)) Then
            If (CInt(Header) > 0) And (CInt(Header) <= Application.Worksheets.Count) Then
                If Sheets(CInt(Header)).Visible = True Then
                    Sheets(CInt(Header)).Select
                    Range("$1:$5").Copy Sheets(W).Range("A1")
                    Sheets(W).Select
                    
                    Range("A1:D4").Interior.ColorIndex = 22
                    Range("E1:E4").Interior.ColorIndex = 40
                    Range("F1:O4").Interior.ColorIndex = 39
                    Range("P1:V4").Interior.ColorIndex = 37
                    Range("W1:AY5").Interior.ColorIndex = -4142
                    
                End If
            End If
        End If
        
        Sheets(Q).Select
        
        For i = 1 To 7
            Columns(1).Insert
        Next i

        Columns(InvoiceCol + 7).EntireColumn.Cut Sheets(Q).Columns(1)
        Columns(DocTypeCol + 7).EntireColumn.Cut Sheets(Q).Columns(2)
        Columns(BU3Col + 7).EntireColumn.Cut Sheets(Q).Columns(3)
        Columns(OpenCol + 7).EntireColumn.Cut Sheets(Q).Columns(4)
        Columns(BU5Col + 7).EntireColumn.Cut Sheets(Q).Columns(6)
        Columns(AccountCol + 7).EntireColumn.Cut Sheets(Q).Columns(7)

        Columns("D:D").EntireColumn.NumberFormat = "#,##0.00"
        Columns("H:BB").EntireColumn.Select
        Columns("H:BB").EntireColumn.Delete

        Columns("A:BB").EntireColumn.AutoFit

        For i = 1 To 15
            Columns(1).Insert
        Next i

        Cells(Rows.Count, 16).End(xlUp).EntireRow.Select
        Range(Selection, Range("1:1")).Select

        For Each Row In Selection.Rows

            Row.Cells(1).NumberFormat = "@"
            Row.Cells(1).Value = "00901"
            Row.Cells(2).Value = "I"
            Row.Cells(3).Value = "9"
            Row.Cells(4).Value = "1"
            Row.Cells(6).Value = Month(GLD)
            Row.Cells(7).Value = Day(GLD)
            Row.Cells(8).Value = Right(Year(GLD), 2)
            Row.Cells(9).Value = Left(Year(GLD), 2)
            Row.Cells(10).Value = Row.Cells(22).Value
            Row.Cells(11).Value = Month(GLD)
            Row.Cells(12).Value = Day(GLD)
            Row.Cells(13).Value = Right(Year(GLD), 2)
            Row.Cells(14).Value = Left(Year(GLD), 2)
            
            Row.Cells(15).NumberFormat = "@"
            
            If CInt(Row.Cells(21).Value) = 1 Or CInt(Row.Cells(21).Value) = 901 Then
                Row.Cells(15).Value = GL(Cycle(CycleVar))
                ChangeCycle
            Else
                If GL.Exists(CInt(Row.Cells(21).Value)) Then
                    Row.Cells(15).Value = GL(CInt(Row.Cells(21).Value))
                End If
            End If
            
            Row.Cells(20).Value = "0"

            For i = 1 To 4
                Row.Cells(i).Interior.ColorIndex = 22
            Next i

            Row.Cells(5).Interior.ColorIndex = 40

            For i = 6 To 15
                Row.Cells(i).Interior.ColorIndex = 39
            Next i

            For i = 16 To 22
                Row.Cells(i).Interior.ColorIndex = 37
            Next i

        Next Row

        Range(Columns(1), Columns(22)).ColumnWidth = 8.43

        For i = 1 To 5

            Range("A1").EntireRow.Insert

        Next i
        
        Sheets(Q).Select
        
        Header = InputBox("Enter Tab # where AUTOCASH header is located.", "Insert Header?", DetailTab + 1)
    
        If (Header <> "") And (IsNumeric(Header)) Then
            If (CInt(Header) > 0) And (CInt(Header) <= Application.Worksheets.Count) Then
                If Sheets(CInt(Header)).Visible = True Then
                    Sheets(CInt(Header)).Select
                    Range("$1:$5").Copy Sheets(Q).Range("A1")
                    Sheets(Q).Select
                    
                    Range("A1:D4").Interior.ColorIndex = 22
                    Range("E1:E4").Interior.ColorIndex = 40
                    Range("F1:O4").Interior.ColorIndex = 39
                    Range("P1:V4").Interior.ColorIndex = 37
                    Range("W1:AY5").Interior.ColorIndex = -4142
                    
                End If
            End If
        End If
    
        For i = 1 To Application.Worksheets.Count
        
            Sheets(i).Select
            Sheets(i).Range("A1").Select
    
        Next i
        
        Sheets(W).Select
        
    End If
    
    Application.StatusBar = False
    
End Sub
