Attribute VB_Name = "RefundJE"
Public Typ As Integer
Public NName As Integer
Public Gross As Integer
Public Store As Integer
Public GL As Integer
Public US As Integer
Public Abort As Boolean

Public DTab As Integer

Public Sub RefundJournalEntry()

    Abort = False

    RefundForm.Show vbModal

    If Abort Then
        Debug.Print ("ABORTED")
        Exit Sub
    End If
    
    Sheets.Add After:=Sheets(Application.Worksheets.Count)
    G = Application.Worksheets.Count
    Sheets.Add After:=Sheets(Application.Worksheets.Count)
    S = Application.Worksheets.Count
    
    Sheets(G).Name = "General Entry"
    Sheets(S).Name = "Summary"
    
    Sheets(G).Select
    
    With Range("A4:B8")
    
        .Interior.ColorIndex = 5
        .Font.ColorIndex = 2
        .Cells(1, 1).Value = "Document Type*"
        .Cells(2, 1).Value = "Explanation*"
        .Cells(3, 1).Value = "GL Date*"
        .Cells(4, 1).Value = "Reversing*"
        .Cells(5, 1).Value = "Memo (BlackLine Only)"
        
        For i = 1 To 5
            .Cells(i, 1).Font.Bold = True
        Next i
    
    End With
    
    With Range("C4:C8")
    
        .Interior.ColorIndex = 36
        .Font.ColorIndex = 1
        .Cells(1).Value = "List - Text:"
        .Cells(2).Value = "List - Text:"
        .Cells(3).Value = "List - Date:"
        .Cells(4).Value = "List - Text:"
        .Cells(5).Value = "List - Text:"
        
    End With
    
    Range("D4").Value = "JE"
    Range("D7").Value = "N"
    Range("H4").Value = "General Entry"
    
    With Range("A4:D8")
        
        .Borders(xlEdgeTop).Weight = -4138
        .Borders(xlEdgeTop).LineStyle = 1
        .Borders(xlEdgeLeft).Weight = -4138
        .Borders(xlEdgeLeft).LineStyle = 1
        .Borders(xlEdgeRight).Weight = -4138
        .Borders(xlEdgeRight).LineStyle = 1
        .Borders(xlEdgeBottom).Weight = -4138
        .Borders(xlEdgeBottom).LineStyle = 1
        
    End With
    
    With Range("A10:D11")
    
        .Borders.Weight = -4138
        .Borders.LineStyle = 1
        
        .Cells(1, 1).Value = "Upl"
        .Cells(1, 2).Value = "Account String*"
        .Cells(1, 3).Value = "Amount*"
        .Cells(1, 4).Value = "Explanation 2"
        .Cells(2, 1).Value = ""
        .Cells(2, 2).Value = "List - Text"
        .Cells(2, 3).Value = "Value"
        .Cells(2, 4).Value = "List - Text"
        
        For i = 1 To 4
            .Cells(1, i).Font.Bold = True
            .Cells(1, i).Interior.ColorIndex = 5
            .Cells(1, i).Font.ColorIndex = 2
            .Cells(2, i).Interior.ColorIndex = 36
            .Cells(2, i).Font.ColorIndex = 1
        Next i
        
    End With
    
    Columns(1).ColumnWidth = 8.43
    Columns(2).ColumnWidth = 14.71
    Columns(3).ColumnWidth = 9.86
    Columns(4).ColumnWidth = 28.43
        
    For i = 5 To 15
        Columns(i).ColumnWidth = 8.43
    Next i
    
    Sheets(S).Select
    
    With Range("A1:D2")
    
        .Cells(1).Value = "Totals By Category"
        .Cells(5).Value = "Goodyear"
        .Cells(6).Value = "Drive"
        .Cells(7).Value = "Goodyear CA"
        .Cells(8).Value = "Drive CA"
        
        .Cells.Font.Bold = True
        .Cells.HorizontalAlignment = xlCenter
        .Cells.Borders.Weight = 2
        
    End With
    
    Range("A1:D1").Merge
    Range("F1:G1").Merge
    Range("F10:G10").Merge
    
    Widths = Array(13.57, 13.57, 13.57, 13.57, 12.57, 13.57, 13.57, 13.57, 13.57, 8.43)
    
    For i = 0 To UBound(Widths)
        Columns(i + 1).ColumnWidth = Widths(i)
    Next i
    
    With Range("F1:G21")
        
        BoldMe = Array(1, 3, 4, 15, 19, 21, 22, 37, 41)
        
        TextCells = Array(1, 3, 4, 5, 7, 9, 11, 13, 15, 19, 21, 22, 23, 25, 27, 31, 33, 35, 37, 41)
        TextMe = Array("Totals By Day", "Day", "AS400 Reports", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Total", "Totals By GL", "GL Acct", "Amount", "901.1110.0308", "901.1110.0318", _
            "901.1150", "907.1110.0308", "907.1110.0318", "907.1150", "Total", "Variance")
        
        LeftMe = Array(3, 4, 5, 7, 9, 11, 13, 21, 22, 23, 25, 27, 31, 33, 35)
        RightMe = Array(15, 37, 41)
            
        For i = 0 To UBound(TextCells)
            .Cells(TextCells(i)).NumberFormat = "@"
            .Cells(TextCells(i)).Value = TextMe(i)
        Next i
        
        For i = 0 To UBound(BoldMe)
            .Cells(BoldMe(i)).Font.Bold = True
        Next i
        
        For i = 6 To 14 Step 2
            .Cells(i).Interior.ColorIndex = 36
        Next i
        
        .Cells(29).Interior.ColorIndex = 1
        .Cells(30).Interior.ColorIndex = 1
        
        .Cells(1).HorizontalAlignment = xlCenter
        .Cells(19).HorizontalAlignment = xlCenter
        
        For i = 0 To UBound(LeftMe)
            .Cells(LeftMe(i)).HorizontalAlignment = xlLeft
        Next i
        
        For i = 0 To UBound(RightMe)
            .Cells(RightMe(i)).HorizontalAlignment = xlRight
        Next i
        
        .Cells(16).Formula = "=SUM(G3:G7)"
        .Cells(38).Formula = "=SUM(G12:G14,G16:G18)"
        .Cells(42).Formula = "=G8+G19"
    
    End With
    
    With Range("H2:I8")
    
        .Cells(1).Value = "GSReports"
        .Cells(2).Value = "Variance"
        
        .Cells(1).Font.Bold = True
        .Cells(2).Font.Bold = True
        
        .Cells(1).HorizontalAlignment = xlLeft
        .Cells(2).HorizontalAlignment = xlLeft
        
        .Cells(13).Formula = "=SUM(H3:H7)"
        .Cells(14).Formula = "=SUM(I3:I7)"
        
        For i = 3 To 12
            .Cells(i).Interior.ColorIndex = 36
        Next i
        
        For i = 3 To 14
            .Cells(i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Next i
        
    End With
    
    Range("F1:G2").Borders.Weight = 2
    Range("H2:I2").Borders.Weight = 2
    Range("F3:I7").Borders(xlEdgeLeft).Weight = 2
    Range("F3:I7").Borders(xlEdgeRight).Weight = 2
    Range("F3:I7").Borders(xlInsideVertical).Weight = 2
    Range("F8:I8").Borders.Weight = 2
    Range("F10:G11").Borders.Weight = 2
    Range("F12:G18").Borders(xlEdgeLeft).Weight = 2
    Range("F12:G18").Borders(xlEdgeRight).Weight = 2
    Range("F12:G18").Borders(xlInsideVertical).Weight = 2
    Range("F19:G19").Borders.Weight = 2
    Range("F21:G21").Borders.Weight = 2
    
    Sheets(DTab).Select
    Range(Range("A1").End(xlDown), Range("1:1")).Select
    
    Dim GAccounts As Collection
    Dim GAmounts As Collection
    Dim GExp As Collection
    
    Dim SGood As Collection
    Dim SDrive As Collection
    Dim SGoodCA As Collection
    Dim SDriveCA As Collection
    
    Dim GLParse() As String
    
    Set GAccounts = New Collection
    Set GAmounts = New Collection
    Set GExp = New Collection
    Set SGood = New Collection
    Set SDrive = New Collection
    Set SGoodCA = New Collection
    Set SDriveCA = New Collection
    
    UseStore = False

    PayType = 0
    
    A9011150 = 0 'BOA
    A9071150 = 0 'BOA CA
    A90111100308 = 0 'GDYR
    A90111100318 = 0 'DRIVE
    A90711100308 = 0 'GDYR CA
    A90711100318 = 0 'DRIVE CA
    
    For Each Row In Selection.Rows
    
        GLParse = Split(Row.Cells(GL).Value, ".")
        
        If UBound(GLParse) = 1 And Not GLParse(0) Like "90#" Then
            UseStore = True
        Else
            UseStore = False
        End If
        
        If UBound(GLParse) > 0 Then
        
            If GLParse(1) = "115" Then
                Row.Cells(GL).NumberFormat = "@"
                Row.Cells(GL).Value = Row.Cells(GL).Value & "0"
            End If
            
        End If
    
        If InStr(1, Row.Cells(Typ).Value, "CHECK", vbTextCompare) > 0 Then
            PayType = 0
        ElseIf InStr(1, Row.Cells(Typ).Value, "VISA", vbTextCompare) > 0 Then
            PayType = 1
        ElseIf InStr(1, Row.Cells(Typ).Value, "MASTER", vbTextCompare) > 0 Then
            PayType = 1
        ElseIf InStr(1, Row.Cells(Typ).Value, "AMERICAN", vbTextCompare) > 0 Then
            PayType = 1
        ElseIf InStr(1, Row.Cells(Typ).Value, "DISCOVER", vbTextCompare) > 0 Then
            PayType = 1
        ElseIf InStr(1, Row.Cells(Typ).Value, "DRIVE", vbTextCompare) > 0 Then
            PayType = 2
        ElseIf InStr(1, Row.Cells(Typ).Value, "GOODYEAR", vbTextCompare) > 0 Then
            PayType = 3
        Else
            PayType = 0
        End If
        
        Select Case PayType
            
            Case 1
                
                If Row.Cells(US).Value = "CA" Then
                    A9071150 = A9071150 - CDbl(Row.Cells(Gross).Value)
                Else
                    A9011150 = A9011150 - CDbl(Row.Cells(Gross).Value)
                End If
                
                If UseStore Then
                    GAccounts.Add (Row.Cells(Store).Value & "." & GLParse(1))
                Else
                    GAccounts.Add (Row.Cells(GL).Value)
                End If
                
                GAmounts.Add (Round(CDbl(Row.Cells(Gross).Value), 2))
                GExp.Add (UCase(Row.Cells(NName).Value) & " " & Row.Cells(Store).Value)
                
            Case 2
            
                If Row.Cells(US).Value = "CA" Then
                    A90711100318 = A90711100318 - CDbl(Row.Cells(Gross).Value)
                    SDriveCA.Add (Row.Cells(Gross).Value)
                Else
                    A90111100318 = A90111100318 - CDbl(Row.Cells(Gross).Value)
                    SDrive.Add (Row.Cells(Gross).Value)
                End If
                
                If UseStore Then
                    GAccounts.Add (Row.Cells(Store).Value & "." & GLParse(1))
                Else
                    GAccounts.Add (Row.Cells(GL).Value)
                End If
                
                GAmounts.Add (Round(CDbl(Row.Cells(Gross).Value), 2))
                GExp.Add (UCase(Row.Cells(NName).Value) & " " & Row.Cells(Store).Value)
                
            Case 3
                
                If Row.Cells(US).Value = "CA" Then
                    A90711100308 = A90711100308 - CDbl(Row.Cells(Gross).Value)
                    SGoodCA.Add (Row.Cells(Gross).Value)
                Else
                    A90111100308 = A90111100308 - CDbl(Row.Cells(Gross).Value)
                    SGood.Add (Row.Cells(Gross).Value)
                End If
                
                If UseStore Then
                    GAccounts.Add (Row.Cells(Store).Value & "." & GLParse(1))
                Else
                    GAccounts.Add (Row.Cells(GL).Value)
                End If
                
                GAmounts.Add (Round(CDbl(Row.Cells(Gross).Value), 2))
                GExp.Add (UCase(Row.Cells(NName).Value) & " " & Row.Cells(Store).Value)
            
        End Select
        
    Next Row
    
    Sheets(S).Select
    SA = 3
    SB = 3
    SC = 3
    SD = 3
    Maxi = 0
    
    For Each Amount In SGood
        Range("A" & SA).Value = Amount
        SA = SA + 1
    Next Amount
    
    For Each Amount In SDrive
        Range("B" & SB).Value = Amount
        SB = SB + 1
    Next Amount
    
    For Each Amount In SGoodCA
        Range("C" & SC).Value = Amount
        SC = SC + 1
    Next Amount
    
    For Each Amount In SDriveCA
        Range("D" & SD).Value = Amount
        SD = SD + 1
    Next Amount
    
    Maxi = WorksheetFunction.Max(SA, SB, SC, SD)
    
    With Range("A" & Maxi & ":D" & Maxi)
    
        For i = 1 To 4
            .Cells(i).Interior.ColorIndex = 36
            .Cells(i).Font.Bold = True
            .Cells(i).Formula = "=SUM(" & Range(Cells(3, i), Cells(Maxi - 1, i)).Address & ")"
            .Cells(i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
        Next i
        
    End With
    
    For i = 1 To 4
    
        Range(Cells(3, i).Address, Cells(Maxi, i).Address).BorderAround Weight:=2
        Range(Cells(Maxi, i).Address).BorderAround Weight:=2
    
    Next i
    
    Range("A3:D" & Maxi).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
    Range("G3:G8").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
    Range("G12:G21").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
    
    Range("A" & Maxi & ":D" & Maxi).Offset(3).Merge (True)
    
    With Range(Range("A" & Maxi).Offset(3, 0), Range("A" & Maxi).Offset(14, 3))
        
        .Cells(1).Value = "NOT PROCESSED"
        .Cells(5).Value = "NAME"
        .Cells(7).Value = "REASON"
        .Cells(8).Value = "AMOUNT"
        
        ToFormat = Array(1, 5, 7, 8)
        
        For i = 0 To UBound(ToFormat)
            .Cells(ToFormat(i)).Font.Bold = True
            .Cells(ToFormat(i)).HorizontalAlignment = xlCenter
            .Cells(ToFormat(i)).Borders.Weight = 2
        Next i
        
        For i = 1 To 8
            .Cells(i).Borders.Weight = 2
        Next i
        
        For i = 9 To 48
        
            .Cells(i).Interior.ColorIndex = 36
            .Cells(i).HorizontalAlignment = xlCenter
            .Cells(i).Borders(xlEdgeRight).Weight = 2
            .Cells(i).Borders(xlEdgeLeft).Weight = 2
            
            If i Mod 4 = 3 Then
                .Cells(i).Font.Italic = True
            ElseIf i Mod 4 = 0 Then
                .Cells(i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End If
            
            If (i - 44) > 0 Then
                .Cells(i).Borders(xlEdgeBottom).Weight = 2
            End If
        
        Next i
        
        .Cells(48).Offset(1, 0).Formula = "=SUM(" & Range(Range("D" & Maxi).Offset(5, 0), Range("D" & Maxi).Offset(14, 0)).Address & ", I8)"
        .Cells(48).Offset(1, 0).Borders.Weight = 2
        .Cells(48).Offset(1, 0).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
        Range(.Cells(9), .Cells(46)).Merge (True)
        Range(.Cells(5), .Cells(6)).Merge (True)
        
    End With
    
    With Range("G12:G18")
    
        .Cells(1).Value = A90111100308
        .Cells(2).Value = A90111100318
        .Cells(3).Value = A9011150
        .Cells(5).Value = A90711100308
        .Cells(6).Value = A90711100318
        .Cells(7).Value = A9071150
        
    End With
    
    Sheets(G).Select
    
    A9011150 = Round(A9011150, 2)
    A9071150 = Round(A9071150, 2)
    A90111100308 = Round(A90111100308, 2)
    A90711100308 = Round(A90711100308, 2)
    A90111100318 = Round(A90111100318, 2)
    A90711100318 = Round(A90711100318, 2)
    
    Dim SVals As Variant
    
    With Range("B12:D17")
    
        SForm = Array(1, 3, 4, 6, 7, 9, 10, 12, 13, 15, 16, 18)
        AForm = Array(2, 5, 8, 11, 14, 17)
        
        SVals = Array("901.1150", A9011150, ".1150 BOA", "907.1150", A9071150, ".1150 BOA", "901.1110.0308", A90111100308, ".0308 GDYR", _
            "907.1110.0308", A90711100308, ".0308 GDYR", "901.1110.0318", A90111100318, ".0318 DRIVE", "907.1110.0318", A90711100318, ".0318 DRIVE")
        
        For i = 0 To UBound(SForm)
            .Cells(SForm(i)).NumberFormat = "@"
        Next i
        
        For i = 0 To UBound(AForm)
            .Cells(AForm(i)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
        Next i
        
        For i = 0 To 17
            .Cells(i + 1).Value = SVals(i)
        Next i
           
    End With
    
    GLineP = 18
    GCol = 2
    
    For i = 17 To 12 Step -1
        
        If Range("C" & i).Value = 0 Then
            Range("C" & i).EntireRow.Delete
            GLineP = GLineP - 1
        End If
    
    Next i
    
    ReviewCount = 0
    Dim ReviewLines As Collection
    Set ReviewLines = New Collection
    
    Dim AllData As Collection
    Set AllData = New Collection
    
    AllData.Add GAccounts
    AllData.Add GAmounts
    AllData.Add GExp
    
    
    For Each Coll In AllData
    
        GLine = GLineP
        
        For Each Entry In Coll
        
            Select Case GCol
                Case 2
                    Range(Cells(GLine, GCol).Address).NumberFormat = "@"
                    Range(Cells(GLine, GCol).Address).Value = Entry
                    
                    If Entry = "901.3016.3" Then
                        Range(Cells(GLine, GCol).Address).Interior.ColorIndex = 3
                        ReviewCount = ReviewCount + 1
                        ReviewLines.Add (Range(Cells(GLine, GCol).Address).Row)
                    End If
                    
                Case 3
                    Range(Cells(GLine, GCol).Address).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
                    Range(Cells(GLine, GCol).Address).Value = Round(Entry, 2)
                Case 4
                    Range(Cells(GLine, GCol).Address).NumberFormat = "@"
                    Range(Cells(GLine, GCol).Address).Value = Entry
            End Select
            
            GLine = GLine + 1
        
        Next Entry
        
        GCol = GCol + 1
        
    Next Coll
    
    With Range("A" & GLine & ":D" & GLine)
    
        .Cells(1).Value = "Totals:"
        .Cells(1).Font.Bold = True
        .Cells(1).HorizontalAlignment = xlLeft
        .Cells(3).Formula = "=SUM(C12:C" & (GLine - 1) & ")"
        .Cells(3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"
        .Cells(3).Font.Bold = True
        .Cells.Borders.Weight = -4138
        .Cells.Borders.LineStyle = 1
        
    End With
    
    Columns("A:D").AutoFit
    Columns("A").ColumnWidth = 8.43
    
    For i = 1 To 4
        Range(Cells(12, i).Address, Cells(GLine, i).Address).BorderAround Weight:=-4138, LineStyle:=1
    Next i
    
    For i = 1 To Application.Worksheets.Count
    
        Select Case Sheets(i).Visible
            Case 0
                Sheets(i).Visible = -1
                Sheets(i).Select
                Range("A1").Select
                Sheets(i).Visible = 0
            Case -1
                Sheets(i).Select
                Range("A1").Select
            Case 2
                Sheets(i).Visible = -1
                Sheets(i).Select
                Range("A1").Select
                Sheets(i).Visible = 2
        End Select
    
    Next i
    
    Sheets(G).Select
    
    ReviewPos = 7
    
    If ReviewCount > 0 Then
        
        Range("H6").Value = "Lines To Review"
    
        For Each Row In ReviewLines
            Range("H" & ReviewPos).Value = Row
            ReviewPos = ReviewPos + 1
        Next Row
        
    End If
    
    Response = MsgBox("JE Lines to Review : " & ReviewCount & vbCr & "JE Totals : " & Sheets(G).Range("C" & GLine).Value, vbOKOnly, "EXEUCTION COMPLETE")
    
    Range("A1").Select
    
End Sub
