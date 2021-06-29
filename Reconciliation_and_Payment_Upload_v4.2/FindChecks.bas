Attribute VB_Name = "FindChecks"
Public Sub FindCheck()

    Dim Upload As Range
    Dim Checks As Range
    Dim Target As Double
    
    Sheets(3).Select
    
    On Error Resume Next
    
        Set Upload = Application.InputBox( _
            Title:="Input Upload Data Region", _
            Prompt:="Input Region of Payment and Sum values - Numbers only. (Usually columns E/F)", _
            Type:=8)
            
        Set Checks = Application.InputBox( _
            Title:="Input Keyed Check Total Region", _
            Prompt:="Input Region of Manually Keyed Checks to run against upload data - Numbers only.", _
            Type:=8)
    
    On Error GoTo 0
    
    If Upload Is Nothing Then
        Sheets(1).Select
        Exit Sub
    End If
    
    If Checks Is Nothing Then
        Sheets(1).Select
        Exit Sub
    End If
    
    Set Upload = Upload.Resize(, Upload.Columns.Count + 1)
    Set Checks = Checks.Resize(, Checks.Columns.Count + 1)
    
    ErrorExists = False
    NotEmpty = False
    WrongSize = False
    
    If Upload.Columns.Count <> 3 Then
        WrongSize = True
    End If
    
    If Checks.Columns.Count <> 2 Then
        WrongSize = True
    End If
    
    For Each Row In Upload.Rows
    
        If IsError(Row.Cells(1)) Or IsError(Row.Cells(2)) Then
            ErrorExists = True
        End If
        
        If Not IsEmpty(Row.Cells(3)) Then
            NotEmpty = True
        End If
        
    Next Row
    
    For Each Row In Checks.Rows
    
        If IsError(Row.Cells(1)) Then
            ErrorExists = True
        End If
        
        If Not IsEmpty(Row.Cells(2)) Then
            NotEmpty = True
        End If
        
    Next Row
    
    If WrongSize Then
        SizeBox = MsgBox("Hey! Your upload selection must be 2 columns. Your check selection must be 1 column only. Try again please.", 16, "Check Yourself")
        Sheets(1).Select
        Exit Sub
    End If
    
    If ErrorExists Then
        ErrorBox = MsgBox("Uh oh! It seems you have included an error in your data selection. Please try again.", 16, "Check Yourself")
        Sheets(1).Select
        Exit Sub
    End If
    
    If NotEmpty Then
        EmptyBox = MsgBox("Hmm... Based on your selection, the adjacent columns to the right still have data in them! Note that this program will overwrite this data. If you would like to keep this data, please exit and adjust the spreadsheet.", 53, "Overwrite Existing Data?")
    End If
    
    If EmptyBox = 2 Then
        Sheets(1).Select
        Exit Sub
    End If

    For Each Row In Upload.Rows
    
        If Row.Cells(2).Value <> "" Then
    
            For Each Check In Checks.Rows
        
                If Row.Cells(3).Value <> "x" Then
            
                    If (Abs(Check.Cells(1).Value - Row.Cells(2).Value) = 0) And Check.Cells(2).Value <> "x" Then
                        Check.Cells(2).Value = "x"
                        
                        Target = Row.Cells(2).Value - Row.Cells(1).Value
                        Row.Cells(3).Value = "x"
                        i = 1
                
                        Do While Target > 0
                
                            Target = Target - Row.Cells(1).Offset(-i, 0).Value
                            Row.Cells(3).Offset(-i, 0).Value = "x"
                            i = i + 1
                    
                        Loop
                        
                        Exit For
                    End If
                    
                End If
            
            Next Check
            
        End If
    
    Next Row
    
    For Each Row In Upload.Rows
    
        If Row.Cells(3).Value <> "x" Then
        
            For Each Check In Checks.Rows
            
                If (Abs(Row.Cells(1).Value - Check.Cells(1).Value) = 0) And Check.Cells(2).Value <> "x" Then
                
                    Check.Cells(2).Value = "x"
                    Row.Cells(3).Value = "x"
                
                    Exit For
                
                End If
            
            Next Check
            
        End If
        
    Next Row
    
    For Each Row In Upload.Rows
        
        If Row.Cells(3).Value <> "x" Then
            Row.Cells(3).Interior.ColorIndex = 3
        End If
        
    Next Row
    
    For Each Check In Checks.Rows
    
        If Check.Cells(2).Value <> "x" Then
            Check.Cells(2).Interior.ColorIndex = 3
        End If
    
    Next Check

End Sub

