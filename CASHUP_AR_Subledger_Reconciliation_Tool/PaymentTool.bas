Attribute VB_Name = "PaymentTool"
Public Task As Integer

Public Abort As Boolean

Public AsOfDate As Date
Public FirstInvoiceDate As Date

Public AccL As Double
Public AccH As Double

Public DateL As Date
Public DateR As Date

Public RPTAB As Integer
Public RPDIVJ As Integer
Public RPDOC As Integer
Public RPAG As Integer
Public RPDCT As Integer
Public RPDCTM As Integer
Public RPGLBA As Integer
Public RPDGJ As Integer

Public Sub PaymentTool()

    For i = Application.Worksheets.Count To 1 Step -1
        
        If Sheets(i).Visible = 0 Then
            Sheets(i).Visible = -1
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Sheets(i).Visible = 0
        End If
        
    Next i
        
    F_PaymentTool.Show
    
    If Abort = True Then
        Exit Sub
    End If
    
    FirstInvoiceDate = CDate("01/01/2100")
    
    Sheets(RPTAB).Select
    
    Sheets(RPTAB).Cells(2, RPDIVJ).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each Cell In Selection.Cells
    
        If CDate(Cell.Value) < FirstInvoiceDate Then
            FirstInvoiceDate = CDate(Cell.Value)
        End If
    
    Next Cell
    
    If Task = 0 Then
    
        Dim TaskNInvs As Dictionary
    
        Set TaskNInvs = New Dictionary
        Set InvBals = New Collection
        Set InvDays = New Collection
        
        Sheets(RPTAB).Cells(1, RPDOC).End(xlDown).Select
        Range(Selection, Range("2:2")).Select
        
        For Each Row In Selection.Rows
        
            If Row.Cells(RPDCTM).Value = "" And (Row.Cells(RPDGJ).Value <= AsOfDate) Then
            
                'Invoice&DocType = [DocType, Gross, GLDate, DaysAfterAsOf]
                TaskNInvs.Add Key:=(Row.Cells(RPDOC).Value & Row.Cells(RPDCT).Value), Item:=Array(Row.Cells(RPDCT).Value, Row.Cells(RPAG).Value, CDate(Row.Cells(RPDGJ).Value), AsOfDate - CDate(Row.Cells(RPDGJ).Value))
                
            End If
            
        Next Row
        
        For Each Row In Selection.Rows
        
            If Row.Cells(RPDCTM).Value <> "" And CDate(Row.Cells(RPDGJ).Value) <= AsOfDate Then
                
                If Row.Cells(RPDCT).Value <> "RU" Or Row.Cells(RPAG).Value >= 0 Or Row.Cells(RPDCTM).Value = "RA" Then
                
                    Prev = TaskNInvs.Item(Row.Cells(RPDOC).Value & Row.Cells(RPDCT).Value)
                    Prev(1) = Row.Cells(RPAG).Value + Prev(1)
                    TaskNInvs.Item(Row.Cells(RPDOC).Value & Row.Cells(RPDCT).Value) = Prev
                
                End If
                
            End If
            
        Next Row
        
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        TN = Application.Worksheets.Count
        Sheets(TN).Name = CStr(Month(AsOfDate) & "." & Day(AsOfDate) & "." & Year(AsOfDate))
        
        Sheets(TN).Select
        
        With Range("A1:E1")
        
            Cells(1).Value = "INVOICE #"
            Cells(2).Value = "INVOICE DATE"
            Cells(3).Value = "DOC TYPE"
            Cells(4).Value = "DAYS OLD"
            Cells(5).Value = "OPEN AMT"
        
        End With
        
        i = 2
        
        For Each Key In TaskNInvs.Keys
        
            If Round(TaskNInvs.Item(Key)(1), 2) <> 0 Then
            
                Range("A" & (i)).Value = Left(Key, Len(Key) - 2)
                Range("B" & (i)).Value = TaskNInvs.Item(Key)(2)
                Range("C" & (i)).Value = TaskNInvs.Item(Key)(0)
                Range("D" & (i)).Value = TaskNInvs.Item(Key)(3)
                Range("E" & (i)).Value = TaskNInvs.Item(Key)(1)
                            
                i = i + 1
                
            End If
            
        Next Key
        
        Range("A:E").Sort Key1:=Range("B:B"), Order1:=xlAscending, Header:=xlYes
        Range("E" & i).Formula = "=SUM(E2:E" & (i - 1) & ")"
        
        Sheets(TN).Range("E:E").Columns.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Sheets(TN).Range("A:E").Columns.AutoFit
            
    
    ElseIf Task = 1 Then
    
        Set TaskIDates = New Collection 'All dates of any activity on acct
        Set TaskIBals = New Collection 'Parallel collection of balances of acct on corresponding date
        
        For i = CDbl(FirstInvoiceDate) To CDbl(Date)
        
            DoEvents
            
            Application.StatusBar = "Calculating balance on DATE : " & CDate(i)
        
            Balance = 0
            
            Sheets(RPTAB).Cells(1, RPDIVJ).End(xlDown).Select
            Range(Selection, Range("2:2")).Select
            
            For Each Row In Selection.Rows
            
                If Row.Cells(RPDCT).Value <> "RU" Or Row.Cells(RPDCTM).Value = "RC" Or Row.Cells(RPDCTM).Value = "RS" Or Row.Cells(RPDCTM).Value = "RA" Then
                
                    If CDbl(CDate(Row.Cells(RPDGJ).Value)) <= i Then
                        Balance = Balance + Row.Cells(RPAG).Value
                    End If
                
                End If
            
            Next Row
            
            If Round(Balance, 2) >= Round(AccL, 2) And Round(Balance, 2) <= Round(AccH, 2) Then
            
                TaskIDates.Add (CDate(i))
                TaskIBals.Add (Round(Balance, 2))
                
            End If
        
        Next i
        
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        TI = Application.Worksheets.Count
        Sheets(TI).Name = Round(AccL, 2) & " < BAL < " & Round(AccH, 2)
        
        Sheets(TI).Select
        
        With Range("A1", "B1")
        
            Cells(1).Value = "DATE"
            Cells(2).Value = "ACCT BALANCE"
        
        End With
        
        For i = 1 To TaskIDates.Count
        
            Range("A" & (i + 1)).Value = TaskIDates.Item(i)
            Range("B" & (i + 1)).Value = TaskIBals.Item(i)
        
        Next i
        
        Sheets(TI).Range("B:B").Columns.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Sheets(TI).Range("A:B").Columns.AutoFit
    
    End If

End Sub

Public Function CHECKSEQ(CK As Range, Target As Range, Base As Integer)

    Valid = False
    
    For i = 1 To Base
    
        NewSum = 0
        
        For ii = 0 To (Base - 1)
        
            NewSum = NewSum + Target.Offset(-Base + i + ii, 0).Value
            
        Next ii
        
        If NewSum = Abs(CK.Value) Then
            Valid = True
        End If
        
        NewSum = 0
        
    Next i
    
    CHECKSEQ = Valid
    
End Function
