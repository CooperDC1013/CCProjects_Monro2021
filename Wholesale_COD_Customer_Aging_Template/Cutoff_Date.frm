VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cutoff_Date 
   Caption         =   "Wholesale COD Customer Aging Report (Friday)"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   OleObjectBlob   =   "Cutoff_Date.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cutoff_Date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AcceptButton_Click()

    Valid = True
    
    v = Validator(monthText.Value, dayText.Value, yearText.Value)
    c = ColumnConverter(accountColumn, invoiceDateColumn, openAmountColumn, docTypeColumn)
    If v = 999 Or c(0) = "!" Or c(1) = "!" Or c(2) = "!" Or c(3) = "!" Then
        Valid = False
    End If
    
    If Valid Then
        tMonth = CInt(monthText.Value)
        tDay = CInt(dayText.Value)
        
        If CInt(yearText.Value) < 100 Then
            y = CInt(yearText.Value) + 2000
        Else
            y = CInt(yearText.Value)
        End If
        
        tYear = y
        
        Match = MatchCredits.Value
        ROA = RoaCredits.Value
        WO = WriteOff.Value
        
        AccountCol = c(0)
        InvCol = c(1)
        OpenCol = c(2)
        DocCol = c(3)
    
        Res = MsgBox("You have entered a Cutoff Date of : " & v, vbSystemModal + vbOKCancel)
        
        If Res = 1 Then
            Cutoff_Date.Hide
        End If
        
    Else
        Res = MsgBox("You fool! Enter valid date and column values please.", vbSystemModal + vbOKOnly)
    End If
    
End Sub

Private Sub BlacklistEdit_Click()
    
    Cutoff_Date.Hide
    BlackList.Show
    
End Sub


Private Sub MatchCredits_Enter()
    Value = Not Value
End Sub

Private Sub RoaCredits_Enter()
    Value = Not Value
End Sub

Public Sub UserForm_Initialize()

    monthText.Value = ""
    dayText.Value = ""
    yearText.Value = ""
    
    MatchCredits.SetFocus
    
    Size = BlackList.BlacklistedAccounts.ListCount - 1
    
    AccountBox.Clear
    
    If Size = -1 Then
        ReDim Friday.BlackListArr(0 To 1, 0 To 0)
        AccountBox.Clear
    Else
        ReDim Preserve Friday.BlackListArr(0 To 1, 0 To Size)
    
        For i = 0 To UBound(Friday.BlackListArr, 2)
            AccountBox.AddItem
            AccountBox.List(i, 0) = Friday.BlackListArr(0, i)
            AccountBox.List(i, 1) = Friday.BlackListArr(1, i)
        Next
    
    End If
    
End Sub

Private Function Validator(month As Variant, day As Variant, year As Variant)

    mbool = True
    dbool = True
    ybool = True
    
    If Not IsNumeric(month) Then
        mbool = False
    End If
    
    If Not IsNumeric(day) Then
        dbool = False
    End If
    
    If Not IsNumeric(year) Then
        ybool = False
    End If
    
    tbool = mbool And dbool And ybool
    
    If tbool Then
        If CInt(year) < 100 Then
            year = CInt(year) + 2000
        End If
        myDate = DateSerial(CInt(year), CInt(month), CInt(day))
    End If
    
    If tbool Then
        Validator = myDate
    Else
        Validator = 999
    End If
    
End Function

Private Function ColumnConverter(Acct As Variant, Inv As Variant, Op As Variant, Doc As Variant)
    
    If IsNumeric(Acct) Then
        Account = CInt(Acct)
    ElseIf IsLetter(Acct) Then
        Account = Range(UCase(Acct) & 1).Column
    Else
        Account = "!"
    End If
    
    If IsNumeric(Inv) Then
        Invoice = CInt(Inv)
    ElseIf IsLetter(Inv) Then
        Invoice = Range(UCase(Inv) & 1).Column
    Else
        Invoice = "!"
    End If
    
    If IsNumeric(Op) Then
        Openn = CInt(Op)
    ElseIf IsLetter(Op) Then
        Openn = Range(UCase(Op) & 1).Column
    Else
        Openn = "!"
    End If
    
    If IsNumeric(Doc) Then
        DocType = CInt(Doc)
    ElseIf IsLetter(Doc) Then
        DocType = Range(UCase(Doc) & 1).Column
    Else
        DocType = "!"
    End If
    
    ColumnConverter = Array(Account, Invoice, Openn, DocType)
    
End Function

Private Function IsLetter(Str)

    For i = 1 To Len(Str)
        
        Select Case Asc(Mid(Str, i, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Friday.Abort = True
    Else
        Friday.Abort = False
    End If
    
End Sub

Private Sub WriteOff_Click()
    Value = Not Value
End Sub

Private Sub WriteOff_Enter()
    Value = Not Value
End Sub
