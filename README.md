<div align="center">

## Create and send Lotus Notes email using COM


</div>

### Description

'

' Creates & sends an email via Lotus Notes 5.0 & up. Also, allows creating/sending even with Lotus Notes not running (although it MUST be loaded on the local machine)
 
### More Info
 
'

' Subject: Subject line of email

' Body: Body (text) of email

' SaveOnSend: True/False, save the email in the 'sent' box

' sendTO (OPTIONAL): the intended receipient of the email

' ccTO (OPTIONAL): the carbon-copied receipient of the email

' bccTO (OPTIONAL): the blind carbon-copied receipient of the email

' lnLogo (OPTIONAL): changes the bitmap logo on the header (0 = no logo)

' AttachmentPath (OPTIONAL): the path of the email attachement

'

' Note: Although all sendto items are optional, if you don't use at least one of them, Lotus Notes won't send your email to any one and may return errors.

'

' Note: Make sure that Lotus Domino Objects is selected as an available reference.

'

' Note: If your Lotus Notes does NOT require a password, then remove the parenthetical ("

----

") following 'Call ses.Initialize'

Returns true if no errors are encountered.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Todd Benson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/todd-benson.md)
**Level**          |Advanced
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/todd-benson-create-and-send-lotus-notes-email-using-com__1-34957/archive/master.zip)





### Source Code

```
Function CreateNewNotesMail(Subject As String, Body As String, SaveOnSend As Boolean, Optional sendTO As String, Optional ccTO As String, Optional bccTO As String, Optional lnLogo As Long, Optional AttachmentPath As String) As Boolean
 Dim ses As New NotesSession  'Notes Session
 Dim mailserver As Variant  'Variable for user's mail server
 Dim mailfile As Variant   'Variable for user's mail file
 Dim lnDatabase As Object  'Notes Database
 Dim lnDocument As Object  'Notes Document
 Dim lnRichText As Object  'Body of Document
 Dim lnAttachment As Object  'Notes Attachement
 On Error GoTo CreateNotesMail_Error
 ' --------------------------------------
 ' Create instantiation of Lotus Notes
 ' Pass Username & password
 ' You can prompt user for password
 ' using inputbox instead of hard coding
 ' password
 ' --------------------------------------
 Call ses.Initialize("*********")  'Replace your email password where the ********* is.
 'Debug.Print ses.UserName
 ' --------------------------------------
 ' Find out the name of the mail server
 ' Find out the name of the mail file
 ' --------------------------------------
 mailserver = ses.GETENVIRONMENTSTRING("Mailserver", True)
 mailfile = ses.GETENVIRONMENTSTRING("Mailfile", True)
 ' --------------------------------------
 ' Open the mail file on the mail server
 ' Create a new email document
 ' --------------------------------------
 Set lnDatabase = ses.GetDatabase(mailserver, mailfile)
 Set lnDocument = lnDatabase.CreateDocument
 Set lnRichText = lnDocument.CreateRichTextItem("Body")
 ' --------------------------------------
 ' Fill out the email text by adding
 ' data passed to the is module
 ' --------------------------------------
 Call lnRichText.AppendText(Body & Chr(13) & Chr(13))
 With lnDocument
  .ReplaceItemValue "SendTo", sendTO
  .ReplaceItemValue "CopyTo", ccTO
  .ReplaceItemValue "BlindCopyTo", bccTO
  .ReplaceItemValue "Subject", Subject
  .ReplaceItemValue "Logo", "StdNotesLtr" & Trim$(str$(lnLogo))
  If SaveOnSend = True Then .SaveMessageOnSend = True
 End With
 ' --------------------------------------
 ' Embed the email attachment, if any
 ' --------------------------------------
 If AttachmentPath <> "" Then
  Set lnAttachment = lnRichText.EMBEDOBJECT(1454, "", AttachmentPath)
 End If
 lnDocument.Send False
 CreateNewNotesMail = True
 ' --------------------------------------
 ' Clean up the code
 ' --------------------------------------
 Set lnDatabase = Nothing
 Set lnDocument = Nothing
 Set lnAttachment = Nothing
CreateNotesMail_Error:
 'Debug.Print Err.Description
 Exit Function
End Function
```

