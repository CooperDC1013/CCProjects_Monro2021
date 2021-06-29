VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HelpScreen 
   Caption         =   "Help Screen"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
   OleObjectBlob   =   "HelpScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HelpScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Page
Public Try

Private Sub Back_Click()
    Try = 0
    
    If Page > 0 Then
        Page = Page - 1
        Update_Help
    End If
End Sub

Private Sub ExitHelp_Click()
    Page = 0
    Unload Me
End Sub

Private Sub Forward_Click()
    
    If Page < 3 Then
        Try = 0
        Page = Page + 1
        Update_Help
    ElseIf Page = 3 Then
        
        Try = Try + 1
        
        If Try = 7 Then
            Page = 4
            Update_Help
        End If
        
    End If
End Sub

Private Sub HelpText_Change()

End Sub

Private Sub UserForm_Activate()
    HelpText.Locked = True
    Page = 0
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub Update_Help()

    If Page = 0 Then
    
        HelpText.Text = vbCr & ":: INTRODUCTION ::" & vbCr & vbCr & _
            "This template allows the user to create a ""Consolidated Invoice"", in which each line item of every" & vbCr & _
            "invoice is listed with its respective 1. Cost of Parts,  2. Cost of Labor, and 3. Description." & vbCr & vbCr & _
            "Each Sub-Invoice also displays the relevant car information as displayed on the original invoice." & vbCr & vbCr & _
            "The Header of the completed document will have the date range, account number, address of" & vbCr & _
            "the customer, and Monro's address as well. Totals will be given here as a" & vbCr & _
            "summary of the consolidated invoice." & vbCr & vbCr & _
            "To run this report, the following tasks must be executed :" & vbCr & vbCr & _
            "1. A query must be ran and dropped into the template with invoice data pulled from the STORE" & vbCr & vbCr & _
            "2. A worldwriter must be ran and dropped into the tempate with invoice data pulled from AR" & vbCr & vbCr & _
            "3. The Pop-Up Form must be keyed with the correct column labels for each field of data" & vbCr & _
            "in each report. This is critical as to ensure accurate transfer of data onto the document." & vbCr & _
            "When the report is created, a new tab will appear in the Excel Workbook that you can print to PDF." & vbCr & _
            "No additional formatting is necessary, just CTRL + P and select PDF File." & vbCr & vbCr & _
            "Please continue to learn how to perform steps 1-3."
            
    ElseIf Page = 1 Then
            
        HelpText.Text = vbCr & ":: STEP 1 STORE QUERY ::" & vbCr & vbCr & _
            "A comprehensive query must be run that takes data from store" & vbCr & _
            "invoices to properly build the report as intended." & vbCr & vbCr & _
            "The query can be run in any format, but must include the following data points :" & vbCr & vbCr & _
            "1. Store #     2. Gross Amount     3. GS/AN Fleet Requirement (OR PO if not Gov't)" & vbCr & _
            "4. Total Tax   5. Total Taxable    6. Make     7. Model    8. Licence Plate" & vbCr & _
            "9. VIN#        10. Mileage" & vbCr & vbCr & _
            "These details are specific to each invoice. The following information must also be included :" & vbCr & vbCr & _
            "11. Item #     12. Item Description   13. Item Parts Price    14. Item Labor Price    15. Qty" & vbCr & vbCr & _
            "The query uses the Qty to multiply out the extended parts and labor prices." & vbCr & _
            "If the Parts/Labor price in the query is already extended, add a column with all 1's for each row." & vbCr & _
            "Use this column as the Qty." & vbCr & vbCr & _
            "Notice that the last items are specific to line item, not invoice, and thus a query is required that" & vbCr & _
            "Separates each invoice by line item, not by whole invoice" & vbCr & vbCr & _
            "Currently, the query CALTRAN exists that fits these parameters." & vbCr & _
            "The files used are F65100, I65101, F655006."
            
    ElseIf Page = 2 Then
    
        HelpText.Text = vbCr & ":: STEP 2 AR WORLDWRITER ::" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & _
            "A basic WW must be run that takes data from the AR account ledger" & vbCr & _
            "to ensure only open invoices are compiled." & vbCr & vbCr & _
            "Any WW can be used, as long as it pulls the Invoice # and Invoice Date for a specific account." & vbCr & _
            "The WW must only display invoices that are currently open/due on account." & vbCr & vbCr & _
            "Typically, a WW that is used for spread adjustments can be used to serve this purpose."
  
    ElseIf Page = 3 Then
    
        HelpText.Text = vbCr & ":: STEP 3 RUNNING THE PROGRAM ::" & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & _
            "Once both of the queries have been dropped as separate tabs in the Excel Workbook," & vbCr & _
            "the previous form may be completed." & vbCr & vbCr & _
            "Enter the proper columnn references (letter names or numbers will work i.e a=1, b=2, c=3...)," & vbCr & _
            "along with the Start and End date of invoices you would like to be included on the report." & vbCr & vbCr & _
            "Then select which logos you would like to appear for specific brands." & vbCr & vbCr & _
            "After your selections are confimed, the report will run for about 10 seconds." & vbCr & _
            "Afterwards, you may print to PDF." & vbCr & vbCr & _
            "You now may exit this screen and proceed with the report."
            
    ElseIf Page = 4 Then
        
        HelpText.Text = vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & _
            "If you are reading this, you are very persistent." & vbCr & vbCr & _
            "But you're also lost." & vbCr & vbCr & _
            "Get outta my office, pleb."
    End If

End Sub

Private Sub UserForm_Initialize()
    Page = 0
    HelpText.Locked = True
End Sub
