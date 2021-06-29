Attribute VB_Name = "Friday"
Public tMonth As Integer 'Var for cutoff month
Public tDay As Integer 'Var for cutoff day
Public tYear As Integer 'Var for cutoff year
Public Match As Boolean 'Var for matching credits checkbox on form
Public ROA As Boolean 'Var for ROA credits checkbox on form
Public WO As Boolean 'Var for Write off checkbox on form
Public AccountCol As Integer 'Var for account number column location on form
Public InvCol As Integer 'Var for invoice date column location on form in format mm/dd/yy
Public OpenCol As Integer 'Var for open amount column location on form
Public DocCol As Integer 'Var for Doctype column location on form
Public BlackListArr() As Variant
Public Abort As Boolean

Public Sub CODAging()
    
    Abort = False

    AccountCol = 1
    InvCol = 1
    OpenCol = 1
    DocCol = 1

    Dim BlackListDict As Dictionary
    Set BlackListDict = Nothing
    Set BlackListDict = New Dictionary
    
    Dim Cutoff As Date
    
    tMonth = 1 'set default cutoff date to 01/01/0101 (literally beginning of time to ensure no invoices are selected if no cutoff is provided.
    tDay = 1
    tYear = 101

    Cutoff_Date.Show vbModal
    
    If Abort = False Then
    
        Cutoff = DateSerial(tYear, tMonth, tDay) 'Serialise the cutoff date into an excel date format
        
        For i = 0 To UBound(BlackListArr, 2)
            
            BlackListDict.Add Key:=Int(BlackListArr(0, i)), Item:=BlackListArr(1, i)
            
        Next
        
        Sheets(1).Name = "MASTER DETAIL"
    
        Dim EverythingStr As String
        
        Rows("1:1").Select 'Activate Filters
        Selection.AutoFilter
        
        Range(Selection, Selection.End(xlToRight)).Select 'Get entire data set
        Range(Selection, Selection.End(xlDown)).Select 'cont'd
        EverythingStr = Selection.Address 'Set Data range as address string for later reference
        
        Sheets.Add After:=ActiveSheet 'Add all relevant sheets and rename
        Sheets.Add After:=ActiveSheet
        Sheets.Add After:=ActiveSheet
        Sheets.Add After:=ActiveSheet
        Sheets(2).Name = "WORKING SHEET"
        Sheets(3).Name = "MINOR WRITE OFFS"
        Sheets(4).Name = "MATCHING CREDITS"
        Sheets(5).Name = "APPLIED CREDITS"
        Sheets(1).Select
        
        If WO = True Then 'If this variable is true then execute Write-Off algorithm below
            ActiveSheet.Range(EverythingStr).AutoFilter Field:=OpenCol, Criteria1:="<1.00" _
                , Operator:=xlAnd, Criteria2:=">-1.00" 'Filter for less than $1
            Cells.Select 'Select all cells visible
            Selection.Copy
            Sheets(3).Select
            Range("A1").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Range("A1").Select
            Sheets(1).Select
            Sheets(1).ShowAllData
        End If
        
        Sheets(1).Select
        Rows(1).Copy Sheets(4).Rows(1)
        Rows(1).Copy Sheets(5).Rows(1)
        Cells.Select
        Selection.Copy
        Sheets(2).Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Rows("1:1").Select 'Activate Filters
        Selection.AutoFilter
        Range(Cells(Rows.Count, AccountCol).Address).End(xlUp).Offset(0, -(AccountCol - 1)).Select 'Find last cell that has data in account number field and offset to column A
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Range("2:2")).Select
        
        TotalRows = Cells(Rows.Count, AccountCol).End(xlUp).row
        StartSection = 2
        EndSection = 5002
        
        Dim MasterDict As New Scripting.Dictionary
        
        Dim Account As Variant
        
        Dim Paster As Double
        Dim PasterROA As Double
        Paster = 2
        PasterROA = 2
        
        Do While StartSection <= TotalRows
        
            DoEvents
            Application.StatusBar = "OPERATING ROWS " & StartSection & "-" & EndSection
            DoEvents
        
            Set MasterDict = New Scripting.Dictionary
            
            Do Until (Range(Cells(EndSection, AccountCol).Address).Value <> Range(Cells(EndSection + 1, AccountCol).Address).Value) Or (Range(Cells(EndSection + 1, AccountCol).Address).Value = Empty)
                EndSection = EndSection + 1
            Loop
            
            Range(StartSection & ":" & EndSection).Select
        
                For Each row In Selection.Rows 'Loop through each row in chunks of 5000
                    
                    InvDate = CDate(row.Cells(InvCol).Value)
                    
                    If InvDate <= Cutoff And (Not BlackListDict.Exists(row.Cells(AccountCol).Value)) Then
                    
                        If Not MasterDict.Exists(row.Cells(AccountCol).Value) Then
                            MasterDict.Add Key:=row.Cells(AccountCol).Value, Item:=New Collection
                        End If
                        
                        MasterDict.Item(row.Cells(AccountCol).Value).Add Array(row.Cells(OpenCol).Value, row.Cells(OpenCol).EntireRow.Address)
                        
                        Dim NewInv(0 To 1) As Variant
                         
                    End If
                    
                Next
            
            If Match = True Then
        
                For Each Account In MasterDict 'Do this process for every account in the Dictionary
                
                    Posi = 1 'Starting position for scrutinizing invoice
                        
                    Do While Posi <= MasterDict(Account).Count 'Scrutinize every invoice and if it's a credit, find a match if possible.
        
                        Target = MasterDict.Item(Account).Item(Posi) 'assign target as wherever Posi is
                        Posi = Posi + 1 'increase posi for next iteration
                
                        If (Target(0) <= -1#) Then 'identify credit
                            
                            Key = True
                            i = 1
                    
                            Do Until (i > MasterDict.Item(Account).Count Or Key = False)
                    
                                If Key = True Then
                                    If TypeName(MasterDict.Item(Account).Item(i)(0)) <> "String" Then
                                        If CDbl(MasterDict.Item(Account).Item(i)(0)) = (Target(0) * (-1)) Then
                                            Key = False
                                
                                            Range(Target(1)).Cut (Sheets(4).Range("A" & Paster))
                                            Range(MasterDict.Item(Account).Item(i)(1)).Cut (Sheets(4).Range("A" & (Paster + 1)))
                                        
                                            Paster = Paster + 2
                                    
                                            MasterDict.Item(Account).Remove (Posi - 1)
                                    
                                            If (Posi - 1) > i Then
                                                Posi = Posi - 2
                                                MasterDict.Item(Account).Remove (i)
                                            ElseIf (Posi - 1) < i Then
                                                Posi = Posi - 1
                                                MasterDict.Item(Account).Remove (i - 1)
                                            End If
                                                
                                        ElseIf Range(MasterDict.Item(Account).Item(i)(1)).Cells(DocCol).Value Like "C#" Then
                                        
                                            Dif = (Target(0) * (-1)) - MasterDict.Item(Account).Item(i)(0)
                                            
                                            j = 1
                                            
                                            If Dif > 0 And (Dif Mod 5 = 0) Then
                                            
                                                AvoidCurrentInv = i
                                            
                                                NSF = True
                                                
                                                Do Until (j > MasterDict.Item(Account).Count Or NSF = False)
                                                    
                                                    If NSF = True Then
                                                        If TypeName(MasterDict.Item(Account).Item(j)(0)) <> "String" Then
                                                            If CDbl(MasterDict.Item(Account).Item(j)(0)) = Dif And (j <> AvoidCurrentInv) Then
                                                                NSF = False
                                                                
                                                                Range(Target(1)).Cut (Sheets(4).Range("A" & Paster))
                                                                Range(MasterDict.Item(Account).Item(i)(1)).Cut (Sheets(4).Range("A" & (Paster + 1)))
                                                                Range(MasterDict.Item(Account).Item(j)(1)).Cut (Sheets(4).Range("A" & (Paster + 2)))
                                                                
                                                                Sheets(4).Range(Paster & ":" & Paster).Interior.Color = vbYellow
                                                                Sheets(4).Range((Paster + 1) & ":" & (Paster + 1)).Interior.Color = vbYellow
                                                                Sheets(4).Range((Paster + 2) & ":" & (Paster + 2)).Interior.Color = vbYellow
                                                                
                                                                Paster = Paster + 3
                                                        
                                                                MasterDict.Item(Account).Remove (Posi - 1)
                                    
                                                                If (Posi - 1) > i Then
                                                                
                                                                    If (Posi - 1) > j Then
                                                                        Posi = Posi - 1
                                                                    End If
                                                                    
                                                                    Posi = Posi - 2
                                                                    MasterDict.Item(Account).Remove (i)
                                                                    
                                                                    If i < j Then
                                                                        MasterDict.Item(Account).Remove (j - 1)
                                                                    Else
                                                                        MasterDict.Item(Account).Remove (j)
                                                                    End If
                                                                    
                                                                ElseIf (Posi - 1) < i Then
                                                                
                                                                    If (Posi - 1) > j Then
                                                                        Posi = Posi - 1
                                                                    End If
                                                                    
                                                                    Posi = Posi - 1
                                                                    MasterDict.Item(Account).Remove (i - 1)
                                                                    
                                                                    If i < j Then
                                                                        MasterDict.Item(Account).Remove (j - 1)
                                                                    Else
                                                                        MasterDict.Item(Account).Remove (j)
                                                                    End If
                                                                    
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                    j = j + 1
                                                
                                                Loop
                                            End If
                                        End If
                                    End If
                                End If
                        
                                i = i + 1
                        
                            Loop
                        End If
                    Loop
                Next
            End If
            
            DoEvents
            
            Dim HangoverAmt As Double
            Dim HangoverRem As Double
            Dim HangoverAdd As String
            HangoverRem = 0#
            
            If ROA = True Then
    
                For Each Account In MasterDict
    
                    Posi = 1
        
                    Do While Posi <= MasterDict(Account).Count
        
                        Target = MasterDict.Item(Account).Item(Posi)
                        Posi = Posi + 1
        
                        If (Target(0) <= -1#) Then
        
                            Remaining = Target(0) * (-1)
                            
                            If HangoverRem <> 0 And HangoverRem <= Remaining Then
                                
                                Range(HangoverAdd).Cut (Sheets(5).Range("A" & PasterROA))
                                HangoverAmt = 0
                                Remaining = Remaining - HangoverRem
                                HangoverRem = 0
                                PasterROA = PasterROA + 1
                            
                            ElseIf HangoverRem <> 0 And HangoverRem > Remaining Then
                                
                                HangoverRem = HangoverRem - Remaining
                                Remaining = 0
                                
                            End If
                            
                            i = 1
        
                            Do While Remaining > 0 And i <= MasterDict.Item(Account).Count
        
                                If i <> (Posi - 1) Then
        
                                    If MasterDict.Item(Account).Item(i)(0) >= 1# And (MasterDict.Item(Account).Item(i)(0) <= Remaining) Then 'Not a credit and less than credit value; apply whole inv
                                        
                                        If Not Range(MasterDict.Item(Account).Item(i)(1)).Cells(DocCol).Value Like "C#" Then
                                            Range(MasterDict.Item(Account).Item(i)(1)).Cut (Sheets(5).Range("A" & PasterROA))
                                            Remaining = Remaining - MasterDict.Item(Account).Item(i)(0)
                                            MasterDict.Item(Account).Remove (i)
                                            
                                            If Posi - 1 > i Then
                                                Posi = Posi - 1
                                            End If
                                            
                                            PasterROA = PasterROA + 1
                                            i = i - 1
                                        End If
            
                                    ElseIf MasterDict.Item(Account).Item(i)(0) >= 1# And (MasterDict.Item(Account).Item(i)(0) > Remaining) Then
                                        
                                        If Not (Range(MasterDict.Item(Account).Item(i)(1)).Cells(DocCol).Value Like "C#") Then 'Not a credit and greater than credit value; apply partial
                                            HangoverAmt = MasterDict.Item(Account).Item(i)(0)
                                            HangoverAdd = MasterDict.Item(Account).Item(i)(1)
                                            HangoverRem = MasterDict.Item(Account).Item(i)(0) - Remaining
                                            
                                            'Previous = MasterDict.Item(Account).Item(i)(0)
                                            'PreviousAdd = MasterDict.Item(Account).Item(i)(1)
                                            'Range(MasterDict.Item(Account).Item(i)(1)).Cells(15).Value = Remaining
                                            'Range(MasterDict.Item(Account).Item(i)(1)).Copy (Sheets(5).Range("A" & PasterROA))
                                            'Range(MasterDict.Item(Account).Item(i)(1)).Cells(15).Value = Previous - Remaining
                                            
                                            MasterDict.Item(Account).Remove (i)
                                            
                                            'If i = MasterDict.Item(Account).Count + 1 Then
                                                'MasterDict.Item(Account).Add Item:=Array(Previous - Remaining, PreviousAdd), After:=i - 1
                                            'Else
                                                'MasterDict.Item(Account).Add Item:=Array(Previous - Remaining, PreviousAdd), Before:=i
                                            'End If
                                            
                                            If Posi - 1 > i Then
                                                Posi = Posi - 1
                                            End If
                                            
                                            Remaining = 0
                                            'PasterROA = PasterROA + 1
                                            i = i - 1
                                        End If
        
                                    End If
                                End If
                                
                                i = i + 1
                                
                            Loop
                        
                            Range(Target(1)).Cells(OpenCol).Value = Range(Target(1)).Cells(OpenCol).Value + Remaining
                            
                            If Remaining = 0 Then
                                Range(Target(1)).Cut (Sheets(5).Range("A" & PasterROA))
                                MasterDict.Item(Account).Remove (Posi - 1)
                                Posi = Posi - 1
                            ElseIf Target(0) + Remaining <> 0 Then
                                Range(Target(1)).Copy (Sheets(5).Range("A" & PasterROA))
                                Range(Target(1)).Cells(OpenCol).Value = -Remaining
                                PreviousAdd = Target(1)
                                MasterDict.Item(Account).Add Item:=Array(-Remaining, PreviousAdd), Before:=(Posi - 1)
                                MasterDict.Item(Account).Remove (Posi)
                                
                            Else
                                Range(Target(1)).Cells(OpenCol).Value = -Remaining
                                PasterROA = PasterROA - 1
                            End If
                            
                            PasterROA = PasterROA + 1
                            
                        End If
                    Loop
                    
                    If HangoverRem <> 0 Then
                        
                        Range(HangoverAdd).Cells(OpenCol).Value = HangoverAmt - HangoverRem
                        Range(HangoverAdd).Copy (Sheets(5).Range("A" & PasterROA))
                        Range(HangoverAdd).Cells(OpenCol).Value = HangoverRem
                        HangoverRem = 0
                        HangoverAmt = 0
                        PasterROA = PasterROA + 1
                        
                    End If
                Next
    
            End If
            
            StartSection = StartSection + 5001
            EndSection = EndSection + 5001
            
            If EndSection > Rows.Count Then
                EndSection = Rows.Count
            End If
            
            Set MasterDict = Nothing
        
        Loop
        
        Range("A1").Select
        
    End If
    
     Application.StatusBar = False
    
End Sub
