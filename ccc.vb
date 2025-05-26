Sub SendEmails()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim customerName As String
    Dim emailAddresses As String
    Dim pdfPath As String
    Dim fso As Object
    Dim templatePath As String
    Dim sentFolder As String
    Dim i As Integer
    Dim customerNote As String
    Dim response As VbMsgBoxResult

    ' Initialize Outlook
    ' Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Set the Excel sheet
    Set ws = ThisWorkbook.Sheets("Customers")
    
    ' Create sent folder if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    sentFolder = ThisWorkbook.Path & "\sent"
    If Not fso.FolderExists(sentFolder) Then
        fso.CreateFolder sentFolder
    End If
    
    ' Define the template path
    ' templatePath = ThisWorkbook.Path & "\maintenance.msg"
    
    ' Get the last row with data in the "Customers" sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each customer
    For i = 2 To lastRow
        customerName = ws.Cells(i, 1).Value
        emailAddresses = ""
        
        ' Get the customer's note from column B
        customerNote = ws.Cells(i, 2).Value
        
        ' Check if the customer has an email address
        emailCol = 3 ' Starting with column C
        Do While ws.Cells(i, emailCol).Value <> ""
            emailAddresses = emailAddresses & ws.Cells(i, emailCol).Value & ";"
            emailCol = emailCol + 1
        Loop
        
        ' Remove the last semicolon
        If Len(emailAddresses) > 0 Then
            emailAddresses = Left(emailAddresses, Len(emailAddresses) - 1)
        End If
        
        ' Set the PDF path for the current customer
        pdfPath = ThisWorkbook.Path & "\" & customerName & ".pdf"
        
        ' Skip if PDF file is missing
        If Not fso.FileExists(pdfPath) Then
            GoTo SkipCustomer
        End If
        
        ' If there's a note, show it and ask if we should skip the customer
        If Len(customerNote) > 0 Then
            response = MsgBox("Customer: " & customerName & vbCrLf & "Note: " & customerNote & vbCrLf & "Do you want to skip this customer?", vbYesNo + vbQuestion, "Customer Note")
            If response = vbYes Then
                GoTo SkipCustomer ' Skip this customer if user chooses "Yes"
            End If
        End If
        
        ' Create email from template
        ' Set MailItem = OutlookApp.CreateItemFromTemplate(templatePath)
        
        ' Set recipient and subject
        ' MailItem.To = emailAddresses
        ' MailItem.Subject = MailItem.Subject & " - " & customerName
        
        ' Attach the PDF
        ' MailItem.Attachments.Add pdfPath
        
        ' Send the email
        ' MailItem.Send
        
        ' Move the PDF to sent folder
        fso.MoveFile pdfPath, sentFolder & "\" & customerName & ".pdf"
        
SkipCustomer:
    Next i
    
    ' Display completion message
    MsgBox "All emails sent successfully.", vbInformation, "Done"
End Sub

' Sub SendEmails()
'    Dim OutlookApp As Object
'    Dim MailItem As Object
'    Dim ws As Worksheet
'    Dim lastRow As Long
'    Dim customerName As String
'    Dim emailAddresses As String
'    Dim pdfPath As String
'    Dim fso As Object
'    Dim templatePath As String
'    Dim sentFolder As String
'    Dim i As Integer

'    ' Initialize Outlook
'    ' Set OutlookApp = CreateObject("Outlook.Application")

'    ' Set the Excel sheet
'    Set ws = ThisWorkbook.Sheets("Customers")

'    ' Create sent folder if it doesn't exist
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    sentFolder = ThisWorkbook.Path & "\sent"
'    If Not fso.FolderExists(sentFolder) Then
'        fso.CreateFolder sentFolder
'    End If

'    ' Define the template path
'    ' templatePath = ThisWorkbook.Path & "\maintenance.msg"

'    ' Get the last row with data in the "Customers" sheet
'    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'    ' Loop through each customer
'    For i = 2 To lastRow
'        customerName = ws.Cells(i, 1).Value
'        emailAddresses = ""

'        ' Loop through columns C, D, E, ... to collect all email addresses
'        emailCol = 3 ' Starting with column C
'        Do While ws.Cells(i, emailCol).Value <> ""
'            emailAddresses = emailAddresses & ws.Cells(i, emailCol).Value & ";"
'            emailCol = emailCol + 1
'        Loop

'        ' Remove the last semicolon
'        If Len(emailAddresses) > 0 Then
'            emailAddresses = Left(emailAddresses, Len(emailAddresses) - 1)
'        End If

'        'customerName = ws.Cells(i, 1).Value
'        'emailAddresses = ws.Cells(i, 3).Value ' Assuming emails are in column C

'        ' Set the PDF path for the current customer
'        pdfPath = ThisWorkbook.Path & "\" & customerName & ".pdf"

'        ' Skip if PDF file is missing
'        If Not fso.FileExists(pdfPath) Then
'            GoTo SkipCustomer
'        End If

'        ' Create email from template
'        ' Set MailItem = OutlookApp.CreateItemFromTemplate(templatePath)

'        ' Set recipient and subject
'        ' MailItem.To = emailAddresses
'        ' MailItem.Subject = MailItem.Subject & " - " & customerName

'        ' Attach the PDF
'        ' MailItem.Attachments.Add pdfPath

'        ' Send the email
'        ' MailItem.Send

'        ' Move the PDF to sent folder
'        fso.MoveFile pdfPath, sentFolder & "\" & customerName & ".pdf"

' SkipCustomer:
'    Next i

'    ' Display completion message
'    MsgBox "All emails sent successfully.", vbInformation, "Done"
' End Sub
