Attribute VB_Name = "AgingAnalysisAlgorithm"
Public RunRF As Boolean
Public MW As Boolean
Public Spread As Boolean
Public RunNSF As Boolean
Public Under10 As Boolean
Public CreditHold As Boolean
Public DocTypeCol As Integer
Public AccountCol As Integer
Public DateCol As Integer
Public InvoiceCol As Integer
Public DueCol As Integer
Public OpenCol As Integer
Public CustAcctCol As Integer
Public TempCreditCol As Integer

Public GrossCol() As Variant
Public BUCol() As Variant
Public BU3Col() As Variant
Public BU5Col() As Variant

Public DetailTab As Integer
Public CustTab As Integer

'M - RF Working sheet
'N - RF Charges to Remove sheet
'Q - Take off CH
'R - Put on CH
'W - Minor Write Off
'S - Spread

Public Sub AgingAnalysis()

    CurrentOp = "Initializing"
    StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp

    For i = Application.Worksheets.Count To 1 Step -1
    
        If Sheets(i).Visible = False Then
            Sheets(i).Visible = True
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Sheets(Application.Worksheets.Count).Visible = False
        End If
        
    Next i
    
    AgingForm.Show vbModal
    
    Application.StatusBar = StatusMess

    RFAddress = 1

    If RunRF = True Then 'Box is checked on UserForm, remove all qualifying RF charges
    
        CurrentOp = "Evaluating RF Charges"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
    
        Sheets.Add After:=Sheets(ActiveWorkbook.Worksheets.Count)
        Sheets(ActiveWorkbook.Worksheets.Count).Name = "RF WORKING SHEET"
        M = ActiveWorkbook.Worksheets.Count
        Sheets.Add After:=Sheets(ActiveWorkbook.Worksheets.Count)
        Sheets(ActiveWorkbook.Worksheets.Count).Name = "RF CHARGES TO REMOVE" 'Create sheet to show all total RF charges present
        N = ActiveWorkbook.Worksheets.Count
        
        Sheets(DetailTab).Range("1:1").Copy Sheets(N).Range("1:1")
        
        Sheets(DetailTab).Select
        Sheets(DetailTab).Name = "MASTER DETAIL"
        Cells.AutoFilter Field:=DocTypeCol, Criteria1:="RF" 'Get separate tab of just RF charges for working with.
        Cells.Select
        Selection.Copy
        Sheets(M).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Sheets(DetailTab).ShowAllData
        Sheets(M).Select
        
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        TotalRows = Cells(Rows.Count, AccountCol).End(xlUp).Row 'Total number of rows is row number of last cell in account column of data set
        StartSection = 2 'Create these variables to use later so dictionary is only full of 2000 accounts to preserve memory.
        EndSection = 2002
        
        Dim RFDict As New Dictionary 'Dict for all RF charges in account
        Dim InvoiceDict As New Dictionary 'Dict for all invoices in account
        Dim PasterRF As Double 'Counter for what row to paste results at
        Paster = 2
        
        Do While StartSection <= TotalRows
            
            Sheets(M).Select
        
            DoEvents
            
            Set RFDict = New Scripting.Dictionary
            Set InvoiceDict = New Scripting.Dictionary
            
            Do Until (Range(Cells(EndSection, AccountCol).Address).Value <> Range(Cells(EndSection + 1, AccountCol).Address).Value) _
                Or (Range(Cells(EndSection + 1, AccountCol).Address).Value = Empty)
                
                EndSection = EndSection + 1
            Loop 'Extend section until last account is not "split" by the end of the section
            
            Range(StartSection & ":" & EndSection).Select 'Select first chunk of accounts
        
            For Each Row In Selection.Rows
            
                If Not RFDict.Exists(Row.Cells(AccountCol).Value) Then 'If the dictionary has no key corresponding to this account
                    RFDict.Add Key:=Row.Cells(AccountCol).Value, Item:=New Collection
                    InvoiceDict.Add Key:=Row.Cells(AccountCol).Value, Item:=New Collection
                End If
                
                If Not Row.Cells(AccountCol).Value = Empty Then
                    RFDict.Item(Row.Cells(AccountCol).Value).Add Array(CDbl(DateValue(Row.Cells(DateCol).Value)), Row.Cells(DateCol).EntireRow.Address)
                End If
                
            Next
            
            Sheets(DetailTab).Select 'Now operating on Master Detail
            Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Range("2:2")).Select
            
            counter = 0
            
            Today = Date
            
            For Each Row In Selection.Rows
            
                If InvoiceDict.Exists(Row.Cells(AccountCol).Value) Then
                    If (Row.Cells(DocTypeCol).Value <> "RF") And (Row.Cells(OpenCol).Value > 0) Then
                    
                        DaysOverdue = Today - (DateValue(Row.Cells(DueCol).Value) + 58) 'Reference is due date; 59 days later will qualify for RF but do 58 later so if today is same as first qualifying RF day, then it is included later.
                        If DaysOverdue > 0 Then
                            counter = counter + 1
                            InvoiceDict.Item(Row.Cells(AccountCol).Value).Add Array(CDbl(DateValue(Row.Cells(DateCol).Value)), CDbl(DaysOverdue), CStr(Row.Cells(DateCol).EntireRow.Address))
                            Debug.Print (Row.Address & " " & counter)
                        End If
                    End If
                End If
                
            Next
            
            'At this point we have the full InvoiceDict of all the invoices for the relevant accounts in the RF section we are operating on. Now we need to work with the due dates
            
            'Section to find accounts with only RFs
            
            For Each Account In RFDict
            
                If InvoiceDict(Account).Count = 0 Then 'No Invoices that have triggered RFs to be created
                    
                    Do While RFDict(Account).Count > 0 'Do while RFs still exist in RFDict
                        
                        Sheets(M).Range(RFDict(Account)(1)(RFAddress)).Cut Sheets(N).Range(Paster & ":" & Paster) 'Paste to next avail line on RFs to Remove Sheet
                        Paster = Paster + 1
                        RFDict(Account).Remove (1) 'Remove RF from RFDict
                        
                    Loop
                    
                    RFDict.Remove (Account) 'Remove Account from RFDict
                    InvoiceDict.Remove (Account) 'Remove Account from InvoiceDict (was empty anyway)
                    
                Else 'There are still overdue invoices, so we will find which one is the oldest and calculate the date to which the oldest RF can exist. Any RFs older will be removed.
                
                    OverdueHighScore = 0
                    
                    For Each Invoice In InvoiceDict(Account)
                    
                        If Invoice(1) > OverdueHighScore Then
                            OverdueHighScore = Invoice(1)
                        End If
                    
                    Next
                    
                    CutoffDate = Today - OverdueHighScore
                    
                    CurrentRF = 1
                    
                    For Each RF In RFDict(Account)
                    
                        If (CDbl(CutoffDate) - RF(0)) >= 0 Then
                        
                            Sheets(M).Range(RF(RFAddress)).Cut Sheets(N).Range(Paster & ":" & Paster)
                            Paster = Paster + 1
                            RFDict(Account).Remove (CurrentRF)
                            CurrentRF = CurrentRF - 1
                        
                        End If
                        
                        CurrentRF = CurrentRF + 1
                    Next
                
                End If
                          
            Next
            
            Set InvoiceDict = Nothing
            Set RFDict = Nothing
            
            StartSection = EndSection + 1
            EndSection = EndSection + 2001
            
        Loop
        
        CurrentOp = "Formatting RF Charges"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
        
        Sheets(DetailTab).Select
        Range("A1").Select
        Sheets(M).Select
        Range("A1").Select
        
        Sheets(N).Select
        
        For i = 1 To 5
            Columns(1).Insert
        Next i
        
        Columns(DocTypeCol + 5).EntireColumn.Cut Sheets(N).Columns(1)
        Columns(AccountCol + 5).EntireColumn.Cut Sheets(N).Columns(2)
        Columns(InvoiceCol + 5).EntireColumn.Cut Sheets(N).Columns(3)
        Columns(DateCol + 5).EntireColumn.Cut Sheets(N).Columns(4)
        Columns(OpenCol + 5).EntireColumn.Cut Sheets(N).Columns(5)

        Columns("E:E").EntireColumn.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
        Columns("F:BB").EntireColumn.Select
        Columns("F:BB").EntireColumn.Delete
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        For Each Cell In Selection
        
            If Cell.Offset(0, 4).Value > 0 Then
                Cell.Value = "LC"
            Else
                Cell.Value = "PL"
                Cell.EntireRow.Font.ColorIndex = 26
            End If
        
        Next Cell
        
        Columns("A:BB").EntireColumn.AutoFit
        Columns("A:E").Select
        Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
        Range("A1").Select
        Columns("A:BB").EntireColumn.AutoFit
    
    End If
    
    Dim Minor As New Dictionary
    
    If MW = True Then
    
        CurrentOp = "Searching For Minor Write Offs"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
    
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        W = Application.Worksheets.Count
        Sheets(W).Name = "MINOR WRITE OFFS"
        PasterW = 2
    
        Sheets(DetailTab).Select
        Range("1:1").Copy Sheets(W).Range("1:1")
        
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            If (Row.Cells(OpenCol).Value < 1#) And (Row.Cells(OpenCol).Value > -1#) And (Row.Cells(DocTypeCol).Value <> "RF") And (Row.Cells(DocTypeCol).Value <> "R5") Then
                    
                If Not Minor.Exists(CDbl(Row.Cells(AccountCol).Value)) Then
                    Minor.Add Key:=CDbl(Row.Cells(AccountCol).Value), Item:=New Dictionary
                End If
                
                Debug.Print (Row.Row & "MW" & PasterW)
                
                Minor(Row.Cells(AccountCol).Value).Add Key:=Row.Cells(InvoiceCol).Value, Item:=Array(Row.Cells(OpenCol).Value, Row.Cells(1).EntireRow.Address)
                Range(CStr(Row.Cells(1).EntireRow.Address)).Copy Sheets(W).Range(PasterW & ":" & PasterW)
                PasterW = PasterW + 1
            
            End If
        
        Next Row
        
        CurrentOp = "Formatting MW AutoCash Template"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
        
        Sheets(W).Select
        
        For i = 1 To 7
            Columns(1).Insert
        Next i
        
        Columns(InvoiceCol + 7).EntireColumn.Cut Sheets(W).Columns(1)
        Columns(DocTypeCol + 7).EntireColumn.Cut Sheets(W).Columns(2)
        
        If BU3Col(0) = True Then
            Columns(BU3Col(1) + 7).EntireColumn.Cut Sheets(W).Columns(3)
        End If
        
        Columns(OpenCol + 7).EntireColumn.Cut Sheets(W).Columns(4)
        
        If BU5Col(0) = True Then
            Columns(BU5Col(1) + 7).EntireColumn.Cut Sheets(W).Columns(6)
        End If
        
        Columns(AccountCol + 7).EntireColumn.Cut Sheets(W).Columns(7)
    
        Columns("D:D").EntireColumn.NumberFormat = "#,##0.00"
        Columns("H:BB").EntireColumn.Select
        Columns("H:BB").EntireColumn.Delete
        
        Columns("A:BB").EntireColumn.AutoFit
        
        For i = 1 To 15
            Columns(1).Insert
        Next i
        
        Vals = Array("RUCO", "RUTRTC", "RUICUT", "RURP3", "BATCH", "G/L MONTH", "G/L DAY", "G/L YEAR", "G/L CENTURY", "CHECK", "CHECK MONTH", "CHECK DAY", "CHECK YEAR", "CHECK CENTURY", "G/L ACCT")
        
        With Range("A1:O1")
        
            For i = 0 To 14
            
                .Cells(i + 1).Value = Vals(i)
                .Cells(i + 1).HorizontalAlignment = xlLeft
            
            Next i
        
        End With
        
        Cells(Rows.Count, 16).End(xlUp).EntireRow.Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            Row.Cells(1).NumberFormat = "@"
            Row.Cells(1).Value = "00901"
            Row.Cells(2).Value = "I"
            Row.Cells(3).Value = "9"
            Row.Cells(4).Value = "1"
            Row.Cells(6).Value = Month(Date)
            Row.Cells(7).Value = Day(Date)
            Row.Cells(8).Value = Right(Year(Date), 2)
            Row.Cells(9).Value = Left(Year(Date), 2)
            Row.Cells(10).Value = Row.Cells(22).Value
            Row.Cells(11).Value = Month(Date)
            Row.Cells(12).Value = Day(Date)
            Row.Cells(13).Value = Right(Year(Date), 2)
            Row.Cells(14).Value = Left(Year(Date), 2)
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
        
        Range(Cells(1, 1), Cells(1, 22)).Interior.ColorIndex = 6
        
        Range(Columns(1), Columns(22)).ColumnWidth = 8.43
        
        For i = 1 To 5
        
            Range("A1").EntireRow.Insert
        
        Next i
        
        'We now have the full list of minor write offs to avoid in the spread in the form {Account # = {Invoice# = [Amt, Address], Invoice# = [...]}, Account# = {...}}
        'Write offs must be accounted for in both the spread and the credit hold section.
        
    End If
    
    Dim SpreadDict As New Dictionary
    Dim AvoidSpread As New Dictionary
    
    If Spread = True Then
    
        CurrentOp = "Searching for Matching Pairs"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
    
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        S = Application.Worksheets.Count
        Sheets(S).Name = "SPREAD UPLOAD"
        PasterS = 2
        
        Sheets(DetailTab).Select
        Range("1:1").Copy Sheets(S).Range("1:1")
        
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        TotalRows = Cells(Rows.Count, AccountCol).End(xlUp).Row 'Total number of rows is row number of last cell in account column of data set
        StartSection = 2 'Create these variables to use later so dictionary is only full of 2000 accounts to preserve memory.
        EndSection = 2002
        
        Do While StartSection <= TotalRows
        
            DoEvents
        
            Do Until (Range(Cells(EndSection, AccountCol).Address).Value <> Range(Cells(EndSection + 1, AccountCol).Address).Value) _
                Or (Range(Cells(EndSection + 1, AccountCol).Address).Value = Empty)
                
                EndSection = EndSection + 1
            Loop 'Extend section until last account is not "split" by the end of the section
            
            Range(StartSection & ":" & EndSection).Select 'Select first chunk of accounts
        
            For Each Row In Selection.Rows
                
                Present = False
            
                If (Row.Cells(DocTypeCol).Value <> "RF") And (Row.Cells(DocTypeCol).Value <> "R5") Then
                
                    If MW = True Then
                        If Minor.Exists(Row.Cells(AccountCol).Value) Then
                            If Minor(Row.Cells(AccountCol).Value).Exists(Row.Cells(InvoiceCol).Value) Then
                                Present = True
                            End If
                        End If
                    End If
                    
                    If Present = False Then
            
                        If Not SpreadDict.Exists(Row.Cells(AccountCol).Value) Then
                            SpreadDict.Add Key:=Row.Cells(AccountCol).Value, Item:=New Dictionary
                        End If
                        
                        If Not SpreadDict(Row.Cells(AccountCol).Value).Exists(Row.Cells(InvoiceCol).Value) Then
                            SpreadDict(Row.Cells(AccountCol).Value).Add Key:=Row.Cells(InvoiceCol).Value, Item:=Array(Row.Cells(OpenCol).Value, Row.Cells(1).EntireRow.Address)
                        End If
                        
                    End If
                        
                End If
            
            Next Row
            
            For Each Account In SpreadDict
            
                For Each Target In SpreadDict(Account)
                
                    Debug.Print (Range(SpreadDict(Account)(Target)(1)).Row)
                
                    If CDbl(SpreadDict(Account)(Target)(0)) <= -1# Then
                    
                        For Each Invoice In SpreadDict(Account)
                        
                            If (CDbl(SpreadDict(Account)(Target)(0)) + CDbl(SpreadDict(Account)(Invoice)(0))) = 0 Then
                            
                                If Not AvoidSpread.Exists(Account) Then
                                    AvoidSpread.Add Key:=Account, Item:=New Dictionary
                                End If
                                
                                AvoidSpread(Account).Add Key:=Target, Item:=Array(SpreadDict(Account)(Target)(0), SpreadDict(Account)(Target)(1))
                                AvoidSpread(Account).Add Key:=Invoice, Item:=Array(SpreadDict(Account)(Invoice)(0), SpreadDict(Account)(Invoice)(1))
                                
                                Range(SpreadDict(Account)(Target)(1)).Copy Sheets(S).Range(PasterS & ":" & PasterS)
                                PasterS = PasterS + 1
                                Range(SpreadDict(Account)(Invoice)(1)).Copy Sheets(S).Range(PasterS & ":" & PasterS)
                                PasterS = PasterS + 1
                                
                                SpreadDict(Account).Remove Key:=Invoice
                                
                                Exit For
                            
                            End If
                        
                        Next Invoice
                    
                    End If
                
                Next Target
            
            Next Account
            
            StartSection = EndSection + 1
            EndSection = EndSection + 2001
            Set SpreadDict = Nothing
            
        Loop
        
        CurrentOp = "Formatting Spread AutoCash Template"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
        
        'The spread sheet is finished and we have the dict AvoidSpread to help for the credit hold portion.
        
        Sheets(S).Select
        
        For i = 1 To 7
            Columns(1).Insert
        Next i
        
        Columns(InvoiceCol + 7).EntireColumn.Cut Sheets(S).Columns(1)
        Columns(DocTypeCol + 7).EntireColumn.Cut Sheets(S).Columns(2)
        
        If BU3Col(0) = True Then
            Columns(BU3Col(1) + 7).EntireColumn.Cut Sheets(S).Columns(3)
        End If
        
        Columns(OpenCol + 7).EntireColumn.Cut Sheets(S).Columns(4)
        
        If BU5Col(0) = True Then
            Columns(BU5Col(1) + 7).EntireColumn.Cut Sheets(S).Columns(6)
        End If
        
        Columns(AccountCol + 7).EntireColumn.Cut Sheets(S).Columns(7)
    
        Columns("D:D").EntireColumn.NumberFormat = "#,##0.00"
        Columns("H:BB").EntireColumn.Select
        Columns("H:BB").EntireColumn.Delete
        
        Columns("A:BB").EntireColumn.AutoFit
        
        For i = 1 To 15
            Columns(1).Insert
        Next i
        
        Vals = Array("RUCO", "RUTRTC", "RUICUT", "RURP3", "BATCH", "G/L MONTH", "G/L DAY", "G/L YEAR", "G/L CENTURY", "CHECK", "CHECK MONTH", "CHECK DAY", "CHECK YEAR", "CHECK CENTURY", "G/L ACCT")
        
        With Range("A1:O1")
        
            For i = 0 To 14
            
                .Cells(i + 1).Value = Vals(i)
                .Cells(i + 1).HorizontalAlignment = xlLeft
            
            Next i
        
        End With
        
        Cells(Rows.Count, 16).End(xlUp).EntireRow.Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            Row.Cells(1).NumberFormat = "@"
            Row.Cells(1).Value = "00901"
            Row.Cells(2).Value = "I"
            Row.Cells(3).Value = "9"
            Row.Cells(4).Value = "1"
            Row.Cells(6).Value = Month(Date)
            Row.Cells(7).Value = Day(Date)
            Row.Cells(8).Value = Right(Year(Date), 2)
            Row.Cells(9).Value = Left(Year(Date), 2)
            Row.Cells(10).Value = Row.Cells(22).Value
            Row.Cells(11).Value = Month(Date)
            Row.Cells(12).Value = Day(Date)
            Row.Cells(13).Value = Right(Year(Date), 2)
            Row.Cells(14).Value = Left(Year(Date), 2)
            Row.Cells(15).NumberFormat = "@"
            Row.Cells(15).Formula = "00040806"
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
        
        Range(Cells(1, 1), Cells(1, 22)).Interior.ColorIndex = 6
        
        Range(Columns(1), Columns(22)).ColumnWidth = 8.43
        
        For i = 1 To 5
        
            Range("A1").EntireRow.Insert
        
        Next i
    
    End If
    
    If CreditHold = True Then 'Detecting all accounts that should not be on credit hold
    
        CurrentOp = "Evaluating Past 60 Day Balances"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
    
        Sheets(CustTab).Name = "CUST DETAIL"
        '===
        Sheets.Add After:=Sheets(ActiveWorkbook.Worksheets.Count)
        Sheets(ActiveWorkbook.Worksheets.Count).Name = "ACCOUNTS TO TAKE OFF HOLD"
        Q = ActiveWorkbook.Worksheets.Count
        Sheets.Add After:=Sheets(ActiveWorkbook.Worksheets.Count)
        Sheets(ActiveWorkbook.Worksheets.Count).Name = "ACCOUNTS TO PUT ON HOLD"
        R = ActiveWorkbook.Worksheets.Count
        
        Sheets(CustTab).Range("1:1").Copy Sheets(Q).Range("1:1")
        Sheets(CustTab).Range("1:1").Copy Sheets(R).Range("1:1")
        '===
        '===
        PasterQ = 2
        PasterR = 2
        
        Cutoff = Date - 30 'Must be less than this to be overdue >60 days
        '===
        '===
        Dim AllAccounts As Scripting.Dictionary
        Set AllAccounts = New Dictionary
        
        '===
        
        Sheets(DetailTab).Select
        
        Cells(2, AccountCol).Select
        Range(Selection, Selection.End(xlDown)).Select
        
        For Each Cell In Selection.Cells
        
            If Not AllAccounts.Exists(Cell.Value) Then
                AllAccounts.Add Key:=Cell.Value, Item:=Array("", "", 0#, 0#)
            End If
            
        Next Cell
        
        Sheets(CustTab).Select
        
        Range(Cells(Rows.Count, CustAcctCol).Address).End(xlUp).Offset(0, -(CustAcctCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            If AllAccounts.Exists(Row.Cells(CustAcctCol).Value) Then
                
                transArray = AllAccounts(Row.Cells(CustAcctCol).Value)
                transArray(0) = CStr(Row.Cells(1).EntireRow.Address)
                transArray(1) = CStr(Row.Cells(TempCreditCol).Value)
                Set AllAccounts(Row.Cells(CustAcctCol).Value) = transArray
                
            End If
        
        Next Row
        
        '===
        
        Dim AvoidRF As Scripting.Dictionary
        Set AvoidRF = New Dictionary
        
        If RunRF = True Then
        
            Sheets(N).Select
            Range("A:E").RemoveSubtotal
            Range("A2").EntireRow.Select
            Range(Selection, Range("A2").End(xlDown).EntireRow).Select
            
            For Each Row In Selection.Rows
            
                AvoidRF.Add Key:=(Row.Cells(2) & Row.Cells(3)), Item:=(Row.Cells(2) & Row.Cells(3))
                '{Acct&Inv = Acct&Inv, ...}
            Next Row
            
            Range("A:E").Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(5), _
            Replace:=True, PageBreaks:=False, SummaryBelowData:=True
            Range("A1").Select
            Columns("A:BB").EntireColumn.AutoFit
        
        End If
        
        '===
        '===
        
        Sheets(DetailTab).Select
        
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            MinorPresent = False
            SpreadPresent = False
        
            If MW = True Then
                If Minor.Exists(Row.Cells(AccountCol).Value) Then
                    If Minor(Row.Cells(AccountCol).Value).Exists(Row.Cells(InvoiceCol).Value) Then
                        MinorPresent = True
                    End If
                End If
            End If
            
            If Spread = True Then
                If AvoidSpread.Exists(Row.Cells(AccountCol).Value) Then
                    If AvoidSpread(Row.Cells(AccountCol).Value).Exists(Row.Cells(InvoiceCol).Value) Then
                        SpreadPresent = True
                    End If
                End If
            End If
        
            If AllAccounts.Exists(Row.Cells(AccountCol).Value) And (SpreadPresent = False) And (MinorPresent = False) Then
        
                If Row.Cells(DocTypeCol).Value = "RF" Then 'Is finance charge
                
                    If Not AvoidRF.Exists(Row.Cells(AccountCol).Value & Row.Cells(InvoiceCol).Value) Then 'Not removed
                    
                        If DateValue(Row.Cells(DueCol).Value) < Cutoff Then
                            
                            transArray = AllAccounts(Row.Cells(AccountCol).Value)
                            transArray(3) = transArray(3) + CDbl(Row.Cells(OpenCol).Value)
                            Set AllAccounts(Row.Cells(AccountCol).Value) = transArray
                            
                        Else
                            
                            transArray = AllAccounts(Row.Cells(AccountCol).Value)
                            transArray(2) = transArray(2) + CDbl(Row.Cells(OpenCol).Value)
                            Set AllAccounts(Row.Cells(AccountCol).Value) = transArray
                            
                        End If
                    End If
                    
                Else
                
                     If DateValue(Row.Cells(DueCol).Value) < Cutoff Then
                            
                            transArray = AllAccounts(Row.Cells(AccountCol).Value)
                            transArray(3) = transArray(3) + CDbl(Row.Cells(OpenCol).Value)
                            Set AllAccounts(Row.Cells(AccountCol).Value) = transArray
                            
                        Else
                            
                            transArray = AllAccounts(Row.Cells(AccountCol).Value)
                            transArray(2) = transArray(2) + CDbl(Row.Cells(OpenCol).Value)
                            Set AllAccounts(Row.Cells(AccountCol).Value) = transArray
                            
                        End If
                End If
                
            End If
            
        Next Row
        '===
        '===
        
        Dim MissingAccounts() As Variant
        X = 0
        
        CurrentOp = "Partitioning Accounts"
        StatusMess = "MACRO IN EXECUTION ; OPERATION : " & CurrentOp
        Application.StatusBar = StatusMess
        
        For Each Account In AllAccounts
        
            If AllAccounts(Account)(1) = "O" Then 'Account is already on hold
                
                If AllAccounts(Account)(3) <= 0 Then '>60 day bal is <= to 0
                    Sheets(CustTab).Range(AllAccounts(Account)(0)).Copy Sheets(Q).Range(PasterQ & ":" & PasterQ)
                    PasterQ = PasterQ + 1
                ElseIf (AllAccounts(Account)(2) + AllAccounts(Account)(3)) <= 0 Then 'Total bal is <= 0
                    Sheets(CustTab).Range(AllAccounts(Account)(0)).Copy Sheets(Q).Range(PasterQ & ":" & PasterQ)
                    PasterQ = PasterQ + 1
                End If
                
            ElseIf AllAccounts(Account)(0) = "" Then 'Acct is missing on all cust
                
                ReDim Preserve MissingAccounts(0 To X)
                MissingAccounts(X) = Account
                X = X + 1
            
            Else 'Account is not on hold
            
                If (AllAccounts(Account)(3) > 0) And ((AllAccounts(Account)(2) + AllAccounts(Account)(3)) > 0) And (AllAccounts(Account)(1) <> "*") Then 'Account should be on hold
                    Sheets(CustTab).Range(AllAccounts(Account)(0)).Copy Sheets(R).Range(PasterR & ":" & PasterR)
                    PasterR = PasterR + 1
                End If
                
            End If
        
        Next Account
        
        Sheets(Q).Select
        Range("A:AA").EntireColumn.AutoFit
        Sheets(R).Select
        Range("A:AA").EntireColumn.AutoFit
    
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        Sheets(Application.Worksheets.Count).Name = "ACCOUNTS MISSING ON CUST DETAIL"
        Y = Application.Worksheets.Count
        Sheets(Y).Select
        Range("A1").Value = "ACCOUNT #"
        PasterY = 2
        
        For i = 0 To UBound(MissingAccounts)
        
            Range("A" & PasterY).Value = MissingAccounts(i)
            PasterY = PasterY + 1
        
        Next i
        
        Columns(1).EntireColumn.AutoFit
    
    End If
    
    For i = 1 To Application.Worksheets.Count
        
        Sheets(i).Select
        Sheets(i).Range("A1").Select
    
    Next i
    
    Application.StatusBar = "EXECUTION SUCCESSFUL"
    
    Sheets(1).Select
    
    Application.StatusBar = False

End Sub
