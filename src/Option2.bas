Attribute VB_Name = "Option2"
Sub SendEmailWithExcelDataAndAttachments()

    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim rngTable As Excel.Range
    Dim sEmailBody As String
    Dim sTemplatePath As String
    Dim sExcelFilePath As String
    Dim sAttachmentPath1 As String
    Dim sAttachmentPath2 As String

    ' Define paths
    sTemplatePath = "C:\YourPath\YourEmailTemplate.oft" ' Replace with your template path
    sExcelFilePath = "C:\YourPath\YourDataFile.xlsx"    ' Replace with your data file path
    sAttachmentPath1 = "C:\YourPath\Attachment1.xlsx"  ' Replace with your attachment path
    sAttachmentPath2 = "C:\YourPath\Attachment2.xlsx"  ' Replace with your attachment path

    On Error GoTo ErrorHandler

    ' Create Outlook Application object
    Set olApp = New Outlook.Application

    ' Create Mail Item from template
    Set olMail = olApp.CreateItemFromTemplate(sTemplatePath)

    ' Open Excel workbook and get data
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(sExcelFilePath)
    Set xlSheet = xlBook.Sheets("Sheet1") ' Replace with your sheet name

    ' Define the range containing the data table
    Set rngTable = xlSheet.Range("A1:D10") ' Adjust range as needed

    ' Copy the range as HTML for better formatting in email
    rngTable.Copy
    sEmailBody = CreateObject("htmlfile").parentWindow.clipboardData.GetData("HTML")

    ' Insert the HTML table into the email body at a placeholder
    ' Assuming your template has a placeholder like "[[TABLE_PLACEHOLDER]]"
    olMail.HTMLBody = Replace(olMail.HTMLBody, "[[TABLE_PLACEHOLDER]]", sEmailBody)

    ' Modify other parts of the email body
    olMail.HTMLBody = Replace(olMail.HTMLBody, "[[NAME]]", "John Doe") ' Example replacement
    olMail.HTMLBody = Replace(olMail.HTMLBody, "[[DATE]]", Format(Date, "dd/mm/yyyy"))

    ' Add attachments
    olMail.Attachments.Add sAttachmentPath1
    olMail.Attachments.Add sAttachmentPath2

    ' Display the email (or use .Send to send directly)
    olMail.Display

ExitHandler:
    ' Clean up
    Set rngTable = Nothing
    If Not xlBook Is Nothing Then xlBook.Close SaveChanges:=False
    Set xlBook = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlApp = Nothing
    Set olMail = Nothing
    Set olApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitHandler

End Sub
