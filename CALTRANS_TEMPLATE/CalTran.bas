Attribute VB_Name = "CalTran"
Public ARInv As Integer
Public ARDate As Integer
Public StoreN As Integer
Public StInv As Integer
Public Gross As Integer
Public PO As Integer
Public Make As Integer
Public Model As Integer
Public Lic As Integer
Public Mile As Integer
Public VIN As Integer
Public Parts As Integer
Public Labor As Integer
Public Taxable As Integer
Public Tax As Integer
Public Qty As Integer
Public ItemN As Integer
Public Desc As Integer
Public Writer As Integer

Public AR As Integer
Public Store As Integer

Public ARHead As Integer
Public StoreHead As Integer

Public Start As Date
Public Finish As Date

Public Confirm As Boolean

Public AllenTire As Boolean
Public Monro As Boolean
Public TireChoice As Boolean
Public MrTire As Boolean
Public TiresNow As Boolean
Public Vac1 As Boolean
Public Vac2 As Boolean
Public Vac3 As Boolean

Private Home As Range

Private Sub LB(Line As Integer)

    With Range("A" & Line, "G" & Line)
        
        .Merge
        .Value = "'==============================================================================================================" '1
        .HorizontalAlignment = xlCenter
    
    End With
    
     Set Home = Range("A" & (Line + 2))
    
End Sub

Private Function Leading(Inp As String)

    If Len(Inp) < 2 And Len(Inp) > 0 Then
        Leading = "0" & Inp
        Exit Function
    ElseIf Len(Inp) >= 2 Then
        Leading = Inp
        Exit Function
    Else
        Leading = "00"
        Exit Function
    End If

End Function

Public Sub CalTran()

    'On Error GoTo NameHandler
    
    For i = 1 To Application.Worksheets.Count

        If Sheets(i).Visible = False Then
            Worksheets(i).Visible = True
            Sheets(i).Move After:=Sheets(Application.Worksheets.Count)
            Worksheets(Application.Worksheets.Count).Visible = False
        End If

    Next i

    CalTranForm2.Show vbModal
    
    'Confirm = True

'    AR = 1
'    Store = 2
'
'    ARInv = 2
'    ARDate = 3
'
'    StoreN = 1
'    StInv = 3
'    Gross = 4
'    PO = 5
'    Make = 7
'    Model = 8
'    Lic = 9
'    Mile = 10
'    VIN = 11
'    Parts = 12
'    Labor = 13
'    Taxable = 14
'    Tax = 15
'    Qty = 16
'    Desc = 17
'    ItemN = 18
    
    Accounting = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    DBA = Array("ALLEN TIRE", "MONRO AUTO SERVICE", "TIRE CHOICE", "MR. TIRE", "TIRES NOW", "CAR-X", "", "")
    
    Dim Brand As String
    
    If AllenTire Then
        Brand = DBA(0)
    ElseIf Monro Then
        Brand = DBA(1)
    ElseIf TireChoice Then
        Brand = DBA(2)
    ElseIf MrTire Then
        Brand = DBA(3)
    ElseIf TiresNow Then
        Brand = DBA(4)
    ElseIf Vac1 Then
        Brand = DBA(5)
    ElseIf Vac2 Then
        Brand = DBA(6)
    ElseIf Vac3 Then
        Brand = DBA(7)
    End If
    
    GrandParts = 0#
    GrandTaxable = 0#
    GrandLabor = 0#
    GrandTax = 0#
    GrandAmt = 0#
    InvoiceCount = 0
    
    If Confirm Then
    
    '''Formatting Header
        Sheets.Add After:=Sheets(Application.Worksheets.Count)
        R = Application.Worksheets.Count
        Sheets(R).Name = "FINISHED DOCUMENT"
        Sheets(R).Cells.Font.Name = "Courier New"
        Sheets(R).Cells.Font.Size = 9
        
        Sheets(R).Columns(1).ColumnWidth = 12.43
        Sheets(R).Columns(2).ColumnWidth = 12.57
        Sheets(R).Columns(3).ColumnWidth = 20.71
        Sheets(R).Columns(4).ColumnWidth = 12.43
        Sheets(R).Columns(5).ColumnWidth = 12.43
        Sheets(R).Columns(6).ColumnWidth = 15.43
        Sheets(R).Columns(7).ColumnWidth = 15.43
        Sheets(R).Columns(8).ColumnWidth = 8.43
        Sheets(R).Columns(9).ColumnWidth = 8.43
        Sheets(R).Columns(10).ColumnWidth = 8.43
        Sheets(R).Columns(11).ColumnWidth = 8.43
        Sheets(R).Columns(12).ColumnWidth = 8.43
        Sheets(R).Columns(13).ColumnWidth = 8.43
        Sheets(R).Columns(14).ColumnWidth = 8.43
        Sheets(R).Columns(15).ColumnWidth = 8.43
        
        With Range("A2", "B3") 'TITLE
            .Merge
            .Value = "INVOICE"
            .Font.Bold = True
            .Font.Size = 26
            .VerticalAlignment = xlCenter
        End With
        
        With Range("B5", "C8") 'HEADER
            .Cells(1).Value = "FROM DATE"
            .Cells(2).Value = "TO DATE"
            .Cells(3).Value = Start
            .Cells(4).Value = Finish
            .Cells(7).Value = "ACCOUNT NO."
            .Cells(8).Value = 3061378
            
            .Cells(3).HorizontalAlignment = xlLeft
            .Cells(4).HorizontalAlignment = xlLeft
            .Cells(8).HorizontalAlignment = xlLeft
            
        End With
        
        With Range("B10", "C11") 'INVOICE #
        
            .Merge (True)
            .Cells(1).Value = "CONSOLIDATED INVOICE #"
            .Cells(3).Value = "I" & Leading(Month(Start)) & Leading(Day(Start)) & Leading(Month(Finish)) & Leading(Day(Finish))
            .Font.Bold = True
            .Cells(1).HorizontalAlignment = xlLeft
            .Cells(3).HorizontalAlignment = xlLeft
            
        End With
        
        With Range("B13", "C13")
        
            .Cells(1).Value = "INVOICE DATE :"
            .Cells(2).Value = Finish
            .Cells(1).HorizontalAlignment = xlRight
            .Cells(2).HorizontalAlignment = xlLeft
            .Cells(2).NumberFormat = "mm/dd/yyyy"
        
        End With
        
        Range("B15", "C17").Merge (True)
        Range("E15", "F17").Merge (True)
        
        With Range("B14", "E18") 'ADDRESSES
            .Cells(4).Value = "REMIT TO :"
            .Cells(5).Value = "CALTRANS DIVISION OF EQUIPMENT"
            .Cells(8).Value = "MNRO HOLDINGS LLC"
            .Cells(9).Value = "691 SOUTH TUSTIN AVENUE"
            
            
            .Cells(12).Value = "DBA " & Brand
            .Cells(13).Value = "ORANGE, CA 92866-3312"
            .Cells(16).Value = "PO BOX 845602"
            .Cells(20).Value = "BOSTON, MA 02284-5580"
        End With
        
        With Range("C24", "C29") 'TOTALS
            .Cells(1).Value = "TOTAL PARTS :"
            .Cells(1).HorizontalAlignment = xlRight
            .Cells(2).Value = "TOTAL PARTS TAXABLE :"
            .Cells(2).HorizontalAlignment = xlRight
            .Cells(3).Value = "TOTAL LABOR :"
            .Cells(3).HorizontalAlignment = xlRight
            .Cells(4).Value = "TOTAL TAX :"
            .Cells(4).HorizontalAlignment = xlRight
            .Cells(6).Value = "INVOICE TOTAL :"
            .Cells(6).HorizontalAlignment = xlRight
            .Cells(6).Font.Bold = True
        End With
        
        Range("E29").Font.Bold = True
        
        LB (31)
        
      '''Starting Data Analysis Here
      
        Dim InvoiceDict As New Scripting.Dictionary
        Set InvoiceDict = Nothing
        Set InvoiceDict = New Dictionary
      
        Sheets(AR).Select
        Range(Cells(2, ARInv).Address).End(xlDown).Offset(0, (1 - ARInv)).Select
        Range(Selection, Range(ARHead & ":" & ARHead)).Select
        
        For Each Row In Selection.Rows
        
            If Row.Cells(ARDate) <= Finish And Row.Cells(ARDate) >= Start Then
                InvoiceDict.Add Key:=CDbl(Row.Cells(ARInv)), Item:=New Collection
                InvoiceDict(CDbl(Row.Cells(ARInv))).Add CDate(Row.Cells(ARDate)) '1
            End If
        
        Next Row
        
        Sheets(Store).Select
        Range(Cells(2, StInv).Address).End(xlDown).Offset(0, (1 - StInv)).Select
        Range(Selection, Range(StoreHead & ":" & StoreHead)).Select
        
        For Each Row In Selection.Rows
        
            If InvoiceDict.Exists(CDbl(Row.Cells(StInv))) Then
        
                If InvoiceDict(CDbl(Row.Cells(StInv))).Count = 1 Then
            
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CDbl(Row.Cells(StoreN))     '2
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CDbl(Row.Cells(Gross))      '3
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CStr(Row.Cells(PO))         '4
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CDbl(Row.Cells(Tax))        '5
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CDbl(Row.Cells(Taxable))    '6
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CStr(Row.Cells(Make))       '7
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CStr(Row.Cells(Model))      '8
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CStr(Row.Cells(Lic))        '9
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CDbl(Row.Cells(Mile))       '10
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add CStr(Row.Cells(VIN))        '11
                    InvoiceDict(CDbl(Row.Cells(StInv))).Add New Collection              '12
                    
                End If
                
                InvoiceDict(CDbl(Row.Cells(StInv)))(12).Add Array(CStr(Row.Cells(ItemN)), CStr(Row.Cells(Desc)), CDbl(Row.Cells(Parts)), CDbl(Row.Cells(Labor)), CInt(Row.Cells(Qty)), CStr(Row.Cells(Writer)))
                
            End If
        
        Next Row
        
        'Have full dict of every invoice. Now need to create for loop for every key in dictionary and creating entry in document for each.
        
        Sheets(R).Select
        
        For Each invoice In InvoiceDict
        
            PartSum = 0#
            LaborSum = 0#
            i = 0
            
            With Range(Home.Offset(0, 4), Home.Offset(7, 4)).Borders(xlEdgeRight)
            
                .LineStyle = xlContinuous
                .Weight = xlThin
            
            End With
        
            With Range(Home, Home.Offset(7, 6))
            
            'First row values
                .Cells(1).Value = "REFERENCE"
                .Cells(2).Value = "STORE"
                .Cells(3).Value = "DATE OF SERVICE"
                .Cells(6).Value = "GROSS AMT"
                .Cells(7).Value = InvoiceDict(invoice)(3)
                GrandAmt = GrandAmt + InvoiceDict(invoice)(3)
                
            'First row format
                .Cells(1).HorizontalAlignment = xlCenter
                .Cells(2).HorizontalAlignment = xlCenter
                .Cells(3).HorizontalAlignment = xlCenter
                .Cells(6).HorizontalAlignment = xlRight
                .Cells(7).HorizontalAlignment = xlRight
                
                .Cells(6).NumberFormat = "* @_)"
                .Cells(7).NumberFormat = Accounting
                
            'Second row values
                .Cells(8).Value = CDbl(invoice)
                .Cells(9).Value = InvoiceDict(invoice)(2)
                .Cells(10).Value = InvoiceDict(invoice)(1)
                .Cells(13).Value = "TOTAL TAX"
                .Cells(14).Value = InvoiceDict(invoice)(5)
                GrandTax = GrandTax + InvoiceDict(invoice)(5)
                
            'Second row format
                .Cells(8).HorizontalAlignment = xlCenter
                .Cells(9).HorizontalAlignment = xlCenter
                .Cells(10).HorizontalAlignment = xlCenter
                .Cells(13).HorizontalAlignment = xlRight
                .Cells(14).HorizontalAlignment = xlRight
                
                .Cells(13).NumberFormat = "* @_)"
                .Cells(14).NumberFormat = Accounting
                
            'Third row values
                .Cells(20).Value = "PARTS TAXABLE"
                .Cells(21).Value = InvoiceDict(invoice)(6)
                GrandTaxable = GrandTaxable + InvoiceDict(invoice)(6)
                
            'Third row format
                .Cells(20).HorizontalAlignment = xlRight
                .Cells(21).HorizontalAlignment = xlRight
                
                .Cells(20).NumberFormat = "* @_)"
                .Cells(21).NumberFormat = Accounting
                
            'Fourth row values
                .Cells(23).Value = "LICENSE"
                .Cells(24).Value = InvoiceDict(invoice)(9)
                .Cells(25).Value = "MAKE"
                .Cells(26).Value = InvoiceDict(invoice)(7)
                
            'Fourth row format
                .Cells(23).HorizontalAlignment = xlRight
                .Cells(24).HorizontalAlignment = xlLeft
                .Cells(25).HorizontalAlignment = xlRight
                .Cells(26).HorizontalAlignment = xlLeft
                
                .Cells(23).NumberFormat = "@* \:_)"
                .Cells(25).NumberFormat = "@* \:_)"
                
            'Fifth row values
                .Cells(30).Value = "VIN"
                .Cells(31).Value = InvoiceDict(invoice)(11)
                .Cells(32).Value = "MODEL"
                .Cells(33).Value = InvoiceDict(invoice)(8)
                .Cells(34).Value = "TOTAL PARTS"
                .Cells(35).Value = "TOTAL LABOR"
                
            'Fifth row format
                .Cells(30).HorizontalAlignment = xlRight
                .Cells(31).HorizontalAlignment = xlLeft
                .Cells(32).HorizontalAlignment = xlRight
                .Cells(33).HorizontalAlignment = xlLeft
                .Cells(34).HorizontalAlignment = xlRight
                .Cells(35).HorizontalAlignment = xlRight
                
                .Cells(30).NumberFormat = "@* \:_)"
                .Cells(32).NumberFormat = "@* \:_)"
                .Cells(34).NumberFormat = "* @_)"
                .Cells(35).NumberFormat = "* @_)"
                
            'Sixth row values
                .Cells(37).Value = "UNIT"
                
                If InvoiceDict(invoice)(4) = "" Then
                    .Cells(38).Value = "FIELD LEFT BLANK"
                Else
                    .Cells(38).Value = InvoiceDict(invoice)(4)
                End If
                
                .Cells(39).Value = "MILEAGE"
                .Cells(40).Value = InvoiceDict(invoice)(10)
                
            'Calculate Part/Labor totals on the fly
                For Each LineItem In InvoiceDict(invoice)(12)
                
                    PartSum = PartSum + (LineItem(2) * LineItem(4))
                    LaborSum = LaborSum + (LineItem(3) * LineItem(4))
                
                Next LineItem
                
                .Cells(41).Value = PartSum
                .Cells(42).Value = LaborSum
                
                GrandParts = GrandParts + PartSum
                GrandLabor = GrandLabor + LaborSum
                
            'Sixth row format
                .Cells(37).HorizontalAlignment = xlRight
                .Cells(38).HorizontalAlignment = xlLeft
                .Cells(39).HorizontalAlignment = xlRight
                .Cells(40).HorizontalAlignment = xlLeft
                .Cells(41).HorizontalAlignment = xlRight
                .Cells(42).HorizontalAlignment = xlRight
                
                .Cells(37).NumberFormat = "@* \:_)"
                .Cells(39).NumberFormat = "@* \:_)"
                .Cells(41).NumberFormat = Accounting
                .Cells(42).NumberFormat = Accounting
            
            'Seventh row is blank
            'Eighth row values
            
                .Cells(50).Value = "QTY"
                .Cells(51).Value = "ITEM #"
                .Cells(52).Value = "ITEM DESCRIPTION"
                .Cells(53).Value = "UNIT PART"
                .Cells(54).Value = "UNIT LABOR"
                .Cells(55).Value = "EXTENDED PART"
                .Cells(56).Value = "EXTENDED LABOR"
                
            'Eighth row format
                
                .Cells(50).HorizontalAlignment = xlRight
                .Cells(51).HorizontalAlignment = xlLeft
                .Cells(52).HorizontalAlignment = xlLeft
                .Cells(53).HorizontalAlignment = xlRight
                .Cells(54).HorizontalAlignment = xlRight
                .Cells(55).HorizontalAlignment = xlRight
                .Cells(56).HorizontalAlignment = xlRight
                
                .Cells(50).NumberFormat = "* @_)"
                .Cells(53).NumberFormat = "* @_)"
                .Cells(54).NumberFormat = "* @_)"
                .Cells(55).NumberFormat = "* @_)"
                .Cells(56).NumberFormat = "* @_)"
            
            End With
            
            For Each LineItem In InvoiceDict(invoice)(12)
            
                With Range(Home.Offset(i + 8, 0), Home.Offset(i + 8, 6))
                
                    .Cells(1).Value = LineItem(4)
                    .Cells(2).Value = LineItem(0)
                    .Cells(3).Value = LineItem(1)
                    .Cells(4).Value = LineItem(2)
                    .Cells(5).Value = LineItem(3)
                    .Cells(6).Value = LineItem(2) * LineItem(4)
                    .Cells(7).Value = LineItem(3) * LineItem(4)
                    
                    .Cells(1).HorizontalAlignment = xlRight
                    .Cells(2).HorizontalAlignment = xlLeft
                    .Cells(3).HorizontalAlignment = xlLeft
                    .Cells(4).HorizontalAlignment = xlRight
                    .Cells(5).HorizontalAlignment = xlRight
                    .Cells(6).HorizontalAlignment = xlRight
                    .Cells(7).HorizontalAlignment = xlRight
                    
                    .Cells(1).NumberFormat = "* ####_)"
                    .Cells(4).NumberFormat = Accounting
                    .Cells(5).NumberFormat = Accounting
                    .Cells(6).NumberFormat = Accounting
                    .Cells(7).NumberFormat = Accounting
                    
                    .Cells(5).Borders(xlEdgeRight).Weight = xlThin
                    .Cells(5).Borders(xlEdgeRight).LineStyle = xlContinuous
                    
                    If Len(LineItem(5)) > 0 Then 'Service writer is populated
                        
                        Dim SubStrings As Variant
                        
                        SubStrings = Split(LineItem(5), " ") 'Split value into words separated by spaces. Save to SubStrings.
                        
                        For ii = 0 To UBound(SubStrings)
                        
                            If Len(SubStrings(ii)) > 19 Then 'If word is greater than 19 characters
                                
                                Mini = ""
                                j = 1
                                
                                Do Until j > Len(SubStrings(ii)) 'Add a space to SubStrings(ii) after 19 characters.
                                    
                                    Mini = Mini & Mid(SubStrings(ii), j, 19) & " "
                                    j = j + 19
                                    
                                Loop
                                
                                SubStrings(ii) = Mini
                                
                            End If
                        
                        Next ii
                        
                        Maxi = ""
                        
                        For Each SubS In SubStrings 'Re-attach each word/19char partial word into one string separated by spaces.
                        
                            Maxi = Maxi & SubS & " "
                        
                        Next SubS
                        
                        SubStrings = Split(Maxi, " ") 'Split again by spaces.
                        
                        k = 0
                        Insert = 1
                        
                        Do While k <= UBound(SubStrings)
                        
                            Liness = ""
                        
                            Do While Len(Liness & " " & SubStrings(k)) <= 20 'All entries in SubStrings are less than 20 characters. Attach as many words to Liness as possible while staying less than 20 chars.
                                
                                Liness = Liness & " " & SubStrings(k)
                                k = k + 1
                                
                                If k > UBound(SubStrings) Then
                                    Exit Do
                                End If
                                
                            Loop
                            
                            .Cells(3).Offset(Insert, 0).Value = Liness
                            .Cells(3).Offset(Insert, 0).HorizontalAlignment = xlLeft
                            .Cells(5).Offset(Insert, 0).Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Cells(5).Offset(Insert, 0).Borders(xlEdgeRight).Weight = xlThin
                            Insert = Insert + 1
                            i = i + 1
                    
                        Loop 'Repeat until all words are included on a line.
                        
                    End If
                    
                End With
                
                i = i + 1
            
            Next LineItem
            
            InvoiceCount = InvoiceCount + 1
            
            LB (CInt(Home.Offset(i + 9, 0).Row))
            
            If (InvoiceCount - 2) Mod 3 = 0 Then
                LB (CInt(Home.Offset(-1, 0).Row))
                Sheets(R).HPageBreaks.Add Before:=Home.Offset(-2, 0)
            End If
        
        Next invoice
        
        With Range("E24", "E29")
        
            .Cells(1).Value = GrandParts
            .Cells(2).Value = GrandTaxable
            .Cells(3).Value = GrandLabor
            .Cells(4).Value = GrandTax
            .Cells(6).Value = GrandAmt
            
            .Cells(1).NumberFormat = Accounting
            .Cells(2).NumberFormat = Accounting
            .Cells(3).NumberFormat = Accounting
            .Cells(4).NumberFormat = Accounting
            .Cells(6).NumberFormat = Accounting
        
        End With
        
        Set InvoiceDict = Nothing
        Sheets(R).PageSetup.LeftMargin = 18
        Sheets(R).PageSetup.RightMargin = 18
        Sheets(R).PageSetup.BottomMargin = 18
        Sheets(R).PageSetup.TopMargin = 18
        Sheets(R).PageSetup.FooterMargin = 3.6
        ActiveWindow.View = xlPageBreakPreview
        Sheets(R).VPageBreaks(1).DragOff xlToRight, 1
        ActiveWindow.View = xlNormalView
        Sheets(R).PageSetup.RightFooter = "&""Courier New,Regular""&10Page &P of &N"
        
        For i = 1 To Application.Worksheets.Count
        
            If Sheets(i).Name = "LOGOS" Then
            
                Sheets(i).Visible = True
                Sheets(i).Select
                
                If AllenTire Then
                    Range("D8", "G12").Copy Sheets(R).Range("D8", "G12")
                ElseIf TireChoice Then
                    Range("D18", "G22").Copy Sheets(R).Range("D9", "G12")
                ElseIf Monro Then
                    Range("D13", "G17").Copy Sheets(R).Range("D8", "G12")
                ElseIf MrTire Then
                    Range("D23", "G27").Copy Sheets(R).Range("D8", "G12")
                ElseIf TiresNow Then
                    Range("D28", "G32").Copy Sheets(R).Range("D8", "G12")
                ElseIf Vac1 Then
                    Range("D33", "G37").Copy Sheets(R).Range("D8", "G12")
                ElseIf Vac2 Then
                    Range("D38", "G42").Copy Sheets(R).Range("D8", "G12")
                ElseIf Vac3 Then
                    Range("D43", "G47").Copy Sheets(R).Range("D8", "G12")
                End If
                
                Range("D2", "G7").Copy Sheets(R).Range("D2", "G7")
                
                Sheets(i).Visible = False
                Sheets(R).Select
            
            End If
        
        Next i
    
        Sheets(AR).Select
        Sheets(AR).Range("A1").Select
        Sheets(Store).Select
        Sheets(Store).Range("A1").Select
        Sheets(R).Select
        Sheets(R).Range("A2").Select
    
    End If
    
    Exit Sub

NameHandler:

    Target = False
    i = 0

    Do Until Target = True

        i = i + 1

        If Sheets(i).Name = "FINISHED DOCUMENT" Then
            Target = True
        End If

    Loop

    Application.DisplayAlerts = False
    Application.Worksheets(i).Delete
    Application.DisplayAlerts = True

    R = R - 1

    Resume
    
End Sub
