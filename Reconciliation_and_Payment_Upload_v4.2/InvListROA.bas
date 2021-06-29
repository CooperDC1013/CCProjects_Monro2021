Attribute VB_Name = "InvListROA"
Public Sub InvListROA()

    On Error Resume Next

    Sheets(10).Select
    Range("A" & Rows.Count).End(xlUp).Offset(1, 1).Select
    Cells(Range("A1").ListObject.ListRows(Range("A1").ListObject.ListRows.Count).Range.Offset(1, 0).Row, 2).Select
    
    Dim WW As New DataObject
    WW.SetText "temp/cooperupl"
    WW.PutInClipboard

End Sub

Public Sub RemoveSuff()
Attribute RemoveSuff.VB_ProcData.VB_Invoke_Func = "C\n14"

    Res = 0
    
    For i = 1 To Len(ActiveCell.FormulaR1C1)
        
        If IsNumeric(Mid(ActiveCell.FormulaR1C1, i, 1)) Then
            Res = Res + 1
        Else
            Exit For
        End If
    
    Next i

    ActiveCell.FormulaR1C1 = Left(ActiveCell.FormulaR1C1, Res)
    ActiveCell.Offset(1, 0).Select
    
End Sub
