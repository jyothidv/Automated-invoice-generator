VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormInvoice 
   Caption         =   "Invoice Entry Form"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   OleObjectBlob   =   "FormInvoice.frx":0000
End
Attribute VB_Name = "FormInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editindex As Integer
Dim isediting As Boolean
Dim invoicePDFpath As String


Private Sub add_Click()

Dim item As ListItem, itemnbr As Integer
      
      'validating input values
      If description.Value = "" Or hour.Value = "" Or rate.Value = "" Then
          MsgBox " please fill in all fields.", vbExclamation
          Exit Sub
      End If
      
      'checking if editing o adding
      If isediting = False Then
             
             'add new item to listview
             itemnbr = ListView1.ListItems.Count + 1
             Set item = ListView1.ListItems.add(, , itemnbr)
             item.SubItems(1) = description.Value
             item.SubItems(2) = Format(hour.Value, "0.00")
             item.SubItems(3) = Format(rate.Value, "0.00")
             item.SubItems(4) = Format(totalamount.Value, "0.00")
             
        Else
        
           'updating item
            With ListView1.ListItems(editindex)
                 item.SubItems(1) = description.Value
                 item.SubItems(2) = Format(hour.Value, "0.00")
                 item.SubItems(3) = Format(rate.Value, "0.00")
                 item.SubItems(4) = Format(totalamount.Value, "0.00")
            End With
            
            isediting = False
            add.Caption = "add"
        End If
        
    ' clear inputs
    description.Value = ""
    hour.Value = ""
    rate.Value = ""
    totalamount = 0
    description.SetFocus
    
    'calculate grand toatl from listview
    Call updategrandtotal
            
          

End Sub

Private Sub address_Change()

End Sub

Private Sub companyname_Change()

End Sub

Private Sub delete_Click()
'Checking list as item or
If ListView1.ListItems.Count = 0 Then
    MsgBox "List is empty", vbExclamation
    Exit Sub
End If

'Checking list items are selectes or not
If ListView1.SelectedItem Is Nothing Then
   MsgBox "Please select an item to remove.", vbExclamation
   Exit Sub
End If

'Remove selected item
 If MsgBox(" do you want to remove item:" & ListView1.SelectedItem.Index, vbYesNo) = vbNo Then Exit Sub
 selectindex = ListView1.SelectedItem.Index
 ListView1.ListItems.Remove selectedindex
 
 'recalculate total
 Call updategrandtotal
 

End Sub

Private Sub updategrandtotal()
    Dim i As Long
    Dim subtotal As Double, discountamt As Double, taxamt As Double, totalamount As Double
    Dim discountrate As Double, taxrate As Double
    
    'calculate subtotal from listview
    For i = 1 To ListView1.ListItems.Count
       subtotal = subtotal + Val(ListView1.ListItems(i).SubItems(4))
       Next i
       txtsubtotal.Value = Format(subtotal, "0.00")
       
    'read discount and tax rate values
    discountamt = Val(discount.Value)
    taxrate = Val(tax.Value)
    
    'calculate amount
    'discountamt = subtotal * discountrate / 100
    taxamt = (subtotal) * taxrate / 100
    totalamount = subtotal - discountamt + taxamt
    
    'update label
    grandtotal.Value = Format(totalamount, "#,##0.00")
    
    
    
End Sub

Private Sub discount_Change()
 If IsNumeric(discount.Value) = False Then
     discount.Value = ""
 Else
     Call updategrandtotal
 End If
 
End Sub

Private Sub edit_Click()
    'checking if list item is selected or not
       If ListView1.SelectedItem Is Nothing Then
          MsgBox "please select an item to edit", vbExclamation
          Exit Sub
        End If
        
     'capturing index and data
         editindex = ListView1.SelectedItem.Index
         With ListView1.ListItems(editindex)
              description.Value = .SubItems(1)
               hour.Value = .SubItems(2)
               rate.Value = .SubItems(3)
         End With
         
    ' setting the editing option
    isediting = True
    add.Caption = "Update"
         
         
End Sub

Private Sub Frame4_Click()

End Sub

Private Sub geninvoice_Click()
     ' validating input values
          If companyname.Value = "" Or email.Value = "" Or phone.Value = "" Or address.Value = "" Then
              MsgBox "customer details missing. please enter all customer details"
              Exit Sub
         End If
         If ListView1.ListItems.Count <= 0 Then
             MsgBox "Invoice item description missing.please enter item description"
             Exit Sub
        End If
        
    'adding customer details
    Sheet1.Range("g10").Value = companyname.Value
    Sheet1.Range("g11").Value = address.Value
    Sheet1.Range("g12").Value = phone.Value
    Sheet1.Range("g13").Value = email.Value
    Sheet1.Range("d10").Value = invoicenumber.Value
    Sheet1.Range("d11").Value = invoicedate.Value
    
    'clearing existing item data
    Sheet1.Range("c17:g27").ClearContents
    
    'adding list item to the sheet
         Dim i As Integer, startrow As Integer: startrow = 17
         For i = 1 To ListView1.ListItems.Count
             With ListView1.ListItems(i)
                 Sheet1.Range("c" & startrow).Value = .Text
                 Sheet1.Range("d" & startrow).Value = .SubItems(1)
                 Sheet1.Range("e" & startrow).Value = .SubItems(2)
                 Sheet1.Range("f" & startrow).Value = .SubItems(3)
                 Sheet1.Range("g" & startrow).Value = .SubItems(4)
            End With
            startrow = startrow + 1
        Next
        Sheet1.Range("g30").Value = discount.Value
        Sheet1.Range("g31").Value = Val(tax.Value) / 100
        
        
        
    'calling generating pdf
    Call generatePDF
    
    MsgBox "Invoice Generated Successfully"
                 
                    
End Sub
Private Sub generatePDF()

On Error GoTo errHandler
   
   'set export data range
       Dim rngtoexport As Range
       Set rngtoexport = Sheet1.Range("A1:I40")
       
    'checking and creating invoice folder if already not there
       Dim invoicefolder As String
       invoicefolder = ThisWorkbook.Path & "\Invoice"
       If Dir(invoicefolder, vbDirectory) = "" Then MkDir invoicefolder
       
    'creating file names
       Dim filename As String, pdfpath As String
       filename = "invoice_" & Sheet1.Range("D10").Value & "_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
       pdfpath = invoicefolder & "\" & filename
       On Error Resume Next
          Kill pdfpath
       On Error GoTo 0
       invoicePDFpath = pdfpath
       
    'exporting to PDF
       rngtoexport.ExportAsFixedFormat Type:=xlTypePDF, filename:=pdfpath, Quality:=x1QualityStandard
       
       Exit Sub
       
errHandler:
    MsgBox "Error:" & Err.description, vbCritical

End Sub

Private Sub hour_Change()
  If IsNumeric(hour.Value) = False Then
       hour.Value = ""
    Else
       
       If hour.Value = "" Or rate.Value = "" Then
           totalamount.Value = 0
       Else
           totalamount.Value = hour.Value * rate.Value
        End If
    End If
    
    

End Sub


Private Sub rate_Change()
If IsNumeric(rate.Value) = False Then
       rate.Value = ""
Else
       
       If rate.Value = "" Or hour.Value = "" Then
           totalamount.Value = 0
       Else
           totalamount.Value = hour.Value * rate.Value
       End If
End If

End Sub


Private Sub reset_Click()

'clearing customer deatils
     companyname.Value = ""
     email.Value = ""
     phone.Value = ""
     address.Value = ""
     
'clearing item details
     description.Value = ""
     hour.Value = ""
     rate.Value = ""
     totalamount.Value = 0
     
     ListView1.ListItems.Clear
     
     discount.Value = 0
     tax.Value = 0
     txtsubtotal.Value = 0
     grandtotal.Value = 0
     
     companyname.SetFocus
     
'generating next invoice number
      Call getinvoicenumber
     
     
     

End Sub

Private Sub sendemail_Click()
     'validating input values
           If companyname.Value = "" Or email.Value = "" Or phone.Value = "" Or address.Value = "" Then
               MsgBox "Customer details missing.please enter all customer details"
               Exit Sub
           End If
           If ListView1.ListItems.Count <= 0 Then
               MsgBox " invoice item description missing,please enter item description"
               Exit Sub
           End If
           
     'generating PDF
           Call generatePDF
           
    ' sending an invoice
          Call sendvoiceEmail
          
    'updating invoice data
         Dim lastrow As Long
         lastrow = Sheet2.Range("A1").CurrentRegion.Rows.Count
         
         Sheet2.Range("A" & lastrow + 1).Value = invoicedate.Value
         Sheet2.Range("B" & lastrow + 1).Value = invoicenumber.Value
         Sheet2.Range("C" & lastrow + 1).Value = companyname.Value
         Sheet2.Range("D" & lastrow + 1).Value = grandtotal.Value
         
         
         'calling reset button
             Call reset_Click
             
        MsgBox "Invoice successfully generated and sent"
         
         
         
End Sub

Private Sub sendvoiceEmail()
       
       'initializing variable
            Dim outlookapp As Object, outlookmail As Object
            Set outlookapp = CreateObject("outlook.application")
            Set outlookmail = outlookapp.createitem(0)
            
        'preparing an email draft
            With outlookmail
               .to = Sheet1.Range("G13").Value
               .Subject = "Invoice from INVOICE"
               .body = "Dear customer," & vbCrLf & vbCrLf & _
                       "Please find attached your invoice." & vbCrLf & vbCrLf & _
                       "Let us know if you have any queries." & vbCrLf & vbCrLf & _
                       "Best reagards," & vbCrLf & " XYZ Pvt L."
               .attachments.add invoicePDFpath
               .Display '.send to send automatically
               '.send
            End With
            Set outlookapp = Nothing
               
End Sub
Private Sub tax_Change()
If IsNumeric(tax.Value) = False Then
     tax.Value = ""
 Else
     Call updategrandtotal
 End If
 
End Sub

Private Sub totalamount_Change()

End Sub

Private Sub UserForm_Initialize()
  invoicedate.Value = Date

  With ListView1
     .View = lvwReport
     .FullRowSelect = True
     .Gridlines = True
     
     .ColumnHeaders.Clear
     .ColumnHeaders.add , , "Item Number", 100
     .ColumnHeaders.add , , "Description", 250
     .ColumnHeaders.add , , "No.of Hours", 102
     .ColumnHeaders.add , , "Hourly Rate", 103
     .ColumnHeaders.add , , "Total Amount", 104
   End With
   Call getinvoicenumber

End Sub

Private Sub getinvoicenumber()
    Dim dtcount As Long
    dtcount = Application.WorksheetFunction.CountIf(Sheet2.Range("A:A"), invoicedate.Value)
    
    Dim invoicenum As String
    invoicenum = "INV-" & Format(invoicedate.Value, "mmddyyyy") & "-" & Format(dtcount + 1, "00000")
    
    invoicenumber.Value = invoicenum
    
End Sub
