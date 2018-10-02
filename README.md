// Put the following Bloack of Codes by creating a new module in Outlook (alt + F11)
//This is to download attachment in Email - work well along with Rules Management in OutLook

Public Sub SaveAttachmentsToDisk(MItem As Outlook.MailItem)
Dim oAttachment As Outlook.Attachment
Dim sSaveFolder As String
sSaveFolder = "C:\Users\shijian.he\Downloads\"
For Each oAttachment In MItem.Attachments
oAttachment.SaveAsFile sSaveFolder & oAttachment.DisplayName
Next
End Sub
