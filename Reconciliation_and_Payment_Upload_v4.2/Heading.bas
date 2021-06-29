Attribute VB_Name = "Heading"
Public Author As String

Public Sub Heading()

    On Error GoTo ErrorHandler

    Title = CStr(ActiveWorkbook.Name)
    TitleF = Replace(Title, ".xlsm", "")
    
    TitleFA = Split(TitleF, " ")
    
    WHSE = CStr(TitleFA(0))
    InpDate = DateValue(Replace(CStr(TitleFA(UBound(TitleFA))), ".", "/"))
    
    Sheets(1).Select
    Range("B4").Value = WHSE
    Range("B5").Value = InpDate
    Range("B6").Value = ""
    
    With Range("B4", "B6")
    
        For i = 1 To 3
        
            .Cells(i).HorizontalAlignment = xlCenter
        
        Next i
    
    End With
    
    Range("B10").Select
    
    Exit Sub
    
ErrorHandler:
ErrorBox = MsgBox("Oops! The title of this document is not in the required format." & vbCr & vbCr & "WHSE #  Name of WHSE  mm/dd/yyyy", 48, "Failure")
Exit Sub

End Sub


