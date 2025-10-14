Attribute VB_Name = "Option1"
   Sub SendEmailWithExcelData()

       Dim olApp As Object 'Outlook Application
       Dim olMail As Object 'Outlook Mail Item
       Dim ws As Worksheet 'Excel Worksheet
       Dim rngData As Range 'Range to copy from Excel
       Dim strEmailBody As String 'Email body content
       Dim strTemplatePath As String 'Path to your email template (optional)
       Dim strAttachmentPath As String 'Path to the file you want to attach

       '--- Configuration ---
       Set ws = ThisWorkbook.Sheets("Sheet1") 'Change "Sheet1" to your sheet name
       Set rngData = ws.Range("A1:C10") 'Change to the actual range you want to copy

       strTemplatePath = "C:\Users\YourUser\Documents\EmailTemplate.oft" 'Optional: Path to your Outlook template (.oft)
       strAttachmentPath = "C:\Users\YourUser\Documents\RelevantFile.xlsx" 'Path to your attachment

       '--- Create Outlook Application and Mail Item ---
       On Error GoTo ErrorHandler
       Set olApp = CreateObject("Outlook.Application")

       If strTemplatePath <> "" Then
           Set olMail = olApp.CreateItemFromTemplate(strTemplatePath) 'Create from template
       Else
           Set olMail = olApp.CreateItem(0) 'Create a new email
       End If

       '--- Prepare Email Body ---
       'Copy data from Excel range to clipboard
       rngData.Copy

       'Get existing email body (if using a template) or start fresh
       strEmailBody = olMail.HTMLBody 'Use .Body for plain text, .HTMLBody for HTML

       'Paste Excel data into the email body (as HTML)
       'This pastes the clipboard content, including formatting
       olMail.HTMLBody = strEmailBody & "<br><br>" & GetClipboardHTML() & "<br><br>"

       '--- Make Modifications to the Email Body ---
       'Example: Add custom text before the pasted data
       olMail.HTMLBody = "Dear Team,<br><br>" & olMail.HTMLBody

       'Example: Add custom text after the pasted data
       olMail.HTMLBody = olMail.HTMLBody & "Kind regards,<br>Your Name"

       '--- Add Attachments ---
       olMail.Attachments.Add strAttachmentPath

       '--- Set Subject and Recipients (customize as needed) ---
       olMail.Subject = "Report from Excel - " & Format(Date, "yyyy-mm-dd")
       olMail.To = "recipient@example.com"
       'olMail.CC = "ccrecipient@example.com"
       'olMail.BCC = "bccrecipient@example.com"

       '--- Display or Send the Email ---
       olMail.Display 'Displays the email for review
       'olMail.Send 'Sends the email without displaying

   Exit Sub

ErrorHandler:
       MsgBox "An error occurred: " & Err.Description, vbCritical
       Set olMail = Nothing
       Set olApp = Nothing

   End Sub

   'Helper function to get HTML content from clipboard
   Private Function GetClipboardHTML() As String
       Dim objData As Object
       Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'DataObject
       objData.GetFromClipboard
       If objData.GetFormat(13) Then '13 = HTML Format
           GetClipboardHTML = objData.GetText(13)
       End If
       Set objData = Nothing
   End Function
