VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BlackList 
   Caption         =   "Blacklist Accounts"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "BlackList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BlackList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddAccount_Click()
    
    BlacklistedAccounts.AddItem
    BlacklistedAccounts.List(BlacklistedAccounts.ListCount - 1, 0) = NewAcctNo.Value
    BlacklistedAccounts.List(BlacklistedAccounts.ListCount - 1, 1) = NewReason.Value
    
    Size = BlacklistedAccounts.ListCount - 1
    
    If Size = -1 Then
        ReDim Friday.BlackListArr(0 To 1, 0 To 0) As Variant
    Else
        ReDim Preserve Friday.BlackListArr(0 To 1, 0 To Size) As Variant
    
        For i = 0 To Size
            Friday.BlackListArr(0, i) = BlacklistedAccounts.List(i, 0)
            Friday.BlackListArr(1, i) = BlacklistedAccounts.List(i, 1)
        Next
    
    End If
    
    NewAcctNo.Value = ""
    NewReason.Value = ""
    NewAcctNo.SetFocus
    
End Sub

Private Sub DeleteAccount_Click()

    Size = BlacklistedAccounts.ListCount - 1
    
    For i = 0 To Size
        If (BlacklistedAccounts.Selected(i) = True) Then
            BlacklistedAccounts.RemoveItem (i)
        End If
    Next
    
    Size = BlacklistedAccounts.ListCount - 1
    
    If Size = -1 Then
        ReDim Friday.BlackListArr(0 To 1, 0 To 0) As Variant
    Else
        ReDim Preserve Friday.BlackListArr(0 To 1, 0 To Size) As Variant
    
        For i = 0 To Size
            Friday.BlackListArr(0, i) = BlacklistedAccounts.List(i, 0)
            Friday.BlackListArr(1, i) = BlacklistedAccounts.List(i, 1)
        Next
        
    End If
        
End Sub

Private Sub ExitButton_Click()
    BlackList.Hide
    Cutoff_Date.UserForm_Initialize
    Cutoff_Date.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    If BlacklistedAccounts.ListCount - 1 = -1 Then
        ReDim Friday.BlackListArr(0 To 1, 0 To 0)
    Else
        ReDim Preserve Friday.BlackListArr(0 To 1, 0 To BlacklistedAccounts.ListCount - 1)
    
        For i = 0 To UBound(Friday.BlackListArr, 2)
            BlacklistedAccounts.AddItem
            BlacklistedAccounts.List(i, 0) = Friday.BlackListArr(0, i)
            BlacklistedAccounts.List(i, 1) = Friday.BlackListArr(1, i)
        Next
    
    End If
    
    NewAcctNo.SetFocus
End Sub
