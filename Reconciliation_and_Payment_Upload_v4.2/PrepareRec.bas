Attribute VB_Name = "PrepareRec"
Sub PrepareRec()
Attribute PrepareRec.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' PrepareRec Macro
'
' Keyboard Shortcut: Ctrl+Shift+C

    If Sheets(1).Range("D5").Value = False Then
        Exit Sub
    End If
    
    Dim Receipts As Boolean
    Dim OrtValid As Boolean
    Dim DatumValid As Boolean
    Dim Ort As String
    Dim Datum As String
    Dim Monat As String
    Dim Tag As String
    Dim OrtHelp As String
    Dim DatumHelp As String
    Dim ConfirmReceipts As String
    
    Dim Comp1 As String
    Dim Comp2 As String
    
    iterator = 0
    
    AddReceipts = MsgBox("Add Unique Receipt ID's to CASH & CC invoices?", (3 + 32), "Prepare Reconciliation for Upload")
    
    Select Case AddReceipts
        Case 2
            Exit Sub
        Case 6
            Receipts = True
        Case 7
            Receipts = False
    End Select
    
    If Receipts Then
        
        OrtValid = (Not IsEmpty(Range("B4"))) And IsNumeric(Range("B4").Value)
        DatumValid = (Not IsEmpty(Range("B5"))) And IsDate(CStr(Range("B5").Value))
        
        If Not (OrtValid And DatumValid) Then
            
            If Not OrtValid Then
                OrtHelp = "Incorrect Location Entry -  Only numeric values accepted."
            Else
                OrtHelp = ""
            End If
            
            If Not DatumValid Then
                DatumHelp = "Incorrect Date Entry - Use only (-) or (/) as delimiter."
            Else
                DatumHelp = ""
            End If
            
            
            KillMsg = MsgBox("The following arguments are invalid :" & vbCr & vbCr & OrtHelp & vbCr & DatumHelp & vbCr & vbCr & "Please update your choices and try again.", vbOKOnly, "Failure")
            
            Exit Sub
        
        End If
        
        Ort = Range("B4").Value
        Datum = Range("B5").Value
        Monat = Month(Datum)
        Tag = Day(Datum)
        
        If Len(Monat) < 2 Then
            Monat = "0" & Monat
        End If
        
        If Len(Tag) < 2 Then
            Tag = "0" & Tag
        End If
        
        ConfirmReceipts = "Unique Receipt Selections :" & vbCr & vbCr & "Location : " & Ort & vbCr & "Date : " & Datum
        
    Else
    
        Ort = ""
        Datum = ""
        Monat = ""
        Tag = ""
        ConfirmReceipts = ""
        
    End If
    
    GoMsg = MsgBox("Warning! Ensure that no #VALUE or #N/A errors exist BETWEEN data entries in any sheet!" & vbCr & vbCr & _
        "#VALUE errors below all data are acceptable." & vbCr & vbCr & vbCr & ConfirmReceipts & vbCr & vbCr & "Select OK to SAVE and CONTINUE.", 49, "Confirm Before Running")
    
    If GoMsg = 2 Then
        Exit Sub
    End If
    
    Application.ActiveWorkbook.Save
    
    ToClean = Array(2, 3, 6, 7, 8, 9)
    
    For Each Sht In ToClean
    
        HideMe = False
    
        Sheets(Sht).Select
        
        Dim Current As Range
        Set Current = Range("A2")
        
        Dim RowsToGo() As Integer
        X = 0
        
        ReDim Preserve RowsToGo(0 To X)
        
        Do Until IsError(Current) Or Current.Row = 2000
            
            If IsEmpty(Current) Then
                ReDim Preserve RowsToGo(0 To X)
                RowsToGo(X) = Current.Row
                X = X + 1
            End If
        
            Set Current = Current.Offset(1, 0)
            
        Loop
        
        If Current.Row = 2 Then
            HideMe = True
        End If
        
        Range(Current, Current.Offset(0, 6)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.ClearContents
        
        Range("A1").Select
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
        
        Target = 1
        
        Do Until Cells(1, Target).Value = "Invoice Raw" Or Target = 20
        
            Target = Target + 1
            
        Loop
        
        Cells(1, Target).EntireColumn.Select
        Selection.Delete
        
        If RowsToGo(0) <> 0 Then
        
            For i = UBound(RowsToGo) To 0 Step -1
            
                Range(RowsToGo(i) & ":" & RowsToGo(i)).Delete
                
            Next i
        
        End If
        
        If (Not HideMe) And Receipts And (Sht <> 3) Then
        
            Columns(2).EntireColumn.Insert
            Cells(1, 2).Value = "Receipt #"
            
            Range(Range("A" & Rows.Count).End(xlUp).Offset(0, 1), "A2").Select
            
            For Each Row In Selection.Rows
            
                CurrentIt = iterator
            
                Do While Len(CurrentIt) < 3
            
                    CurrentIt = "0" & CurrentIt
            
                Loop
                
                If IsError(Row.Cells(1)) Then
                    Comp1 = "000"
                Else
                    Comp1 = CStr(Row.Cells(1).Value)
                End If
                
                If IsError(Row.Cells(1).Offset(-1, 0)) Then
                    Comp2 = "999"
                Else
                    Comp2 = CStr(Row.Cells(1).Offset(-1, 0).Value)
                End If
                
                Identical = (Comp1 = Comp2)
                
                Row.Cells(2).NumberFormat = "@"
                
                If Identical Then
                    Row.Cells(2).Value = Row.Cells(2).Offset(-1, 0).Value
                Else
                    Row.Cells(2).Value = CStr(Monat & Tag & Right(Ort, 1) & CurrentIt)
                    
                    If iterator = 999 Then
                        iterator = 0
                    Else
                        iterator = iterator + 1
                    End If
                End If
            
            Next Row
        
        End If
        
        Range("A1").Select
        
        If HideMe Then
            Sheets(Sht).Visible = 0
        End If
        
        Erase RowsToGo
        
    Next Sht

    Sheets(1).Select
    Range("B4").Select

End Sub
