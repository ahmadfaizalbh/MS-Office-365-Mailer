# MS-Office-365-Mailer
Microsoft Office 365 mailer

### Example code
```python
from office365 import MSOffice365
mailbox = MSOffice365('username@company.com','password','test@company.com')
```
## Read mails
```python
mailbox.Messages(q=query,mail_id=mailId,folder_id=folderId)
```
### Example 1
```python
mails = mailbox.Messages()
list(mails['value'][0])
#[u'IsDeliveryReceiptRequested', u'From', u'HasAttachments', u'WebLink', u'BccRecipients',
# u'ParentFolderId', u'Body', u'Importance', u'ConversationId', u'Categories', u'CcRecipients',
# u'BodyPreview', u'Sender', u'DateTimeCreated', u'@odata.id', u'IsDraft', u'ChangeKey',
# u'DateTimeSent', u'DateTimeReceived', u'DateTimeLastModified', u'ReplyTo', u'ToRecipients', 
# u'IsRead', u'IsReadReceiptRequested', u'@odata.etag', u'Id', u'Subject']
```
### Example 2
```python
mails = mailbox.Messages(q={"Select":["Id","Subject","From"],"top":"50"},folder_id='inbox')
list(mails['value'][0])
#[u'Id', u'From', u'@odata.id', u'Subject', u'@odata.etag']
```
### Example 3
To get a specific mail
```python
mail = mailbox.Messages(q={"Select":["Id","Subject","From"]},mail_id="AGHJS767828KJDS8UJ892WKJSUIKJSK",folder_id='inbox')
list(mail)
#[u'Id', u'From', u'@odata.id', u'Subject', u'@odata.etag']
```

## Send mail
```python
mailbox.Sendmail(
    Subject=mail_Subject,
    Importance=mail_importance,
    Body={
          "ContentType": ContentType,
          "Content": Content
        },
    ToRecipients=[{
          "EmailAddress": {
              "Name": ToDisplayName,
              "Address": ToEmailAdress    
              }
        }],
    Attachments=[list of files],
	SaveToSentItems=True
)
```
### Example
```python
mailbox.Sendmail(
	Subject="Test Mail",
  	Importance="High",
  	Body={
  	    "ContentType": "HTML",
  	    "Content": """<html><body>Hi,<br/>
            <h3>This is new test mail</h3><br/>
            <code>This mail is generated from system </code>
            </body></html>
            """
    },
  	ToRecipients=[{	"EmailAddress": {  "Name": "Test", "Address": "test@company.com" 	}}],
)
```

## Create Draft Message
```python
mailbox.CreateDraftMessage(
	folder_id=folder_id,
    Subject=mail_Subject,
    Importance=mail_importance,
    Body={
          "ContentType": ContentType,
          "Content": Content
        },
    ToRecipients=[{
          "EmailAddress": {
              "Name": ToDisplayName,
              "Address": ToEmailAdress    
              }
        }],
    Attachments=[list of files]
)
```
### Example
```python
mailbox.CreateDraftMessage(
    folder_id='inbox',
	Subject="Test Mail",
  	Importance="High",
  	Body={
  	    "ContentType": "HTML",
  	    "Content": """<html><body>Hi,<br/>
            <h3>This is new test mail</h3><br/>
            <code>This mail is generated from system </code>
            </body></html>
            """
    },
    ToRecipients=[{	"EmailAddress": {  "Name": "Test", "Address": "test@company.com" 	}}],
    Attachments=[`test1.docx','test2.docx'])
```

## Create Folder
```python
mailbox.CreateFolder(folderId, DisplayName)
```
## Create Contact
```python
mailbox.CreateContact(
            GivenName = "Your Name",
            EmailAddresses = [{
                                "Address":"username@company.com",
                                "Name":"Your Name"
                              }],
            BusinessPhones = ["123-456-7890"])
```
## Get Folders
```python
mailbox.Folders(folder_id=folderId, q=query)
```

## Get Calendars
```python
mailbox.Calendars(Calender_id=CalenderId, q=query)
```

## Get Calendar Groups
```python
mailbox.CalendarGroups(CalGroup_id=CalGroupId, q=query)
```

## Get Events
```python
mailbox.Events(Event_id=EventId, q=query)
```

## Get Contacts
```python
mailbox.Contacts(self, Contact_id=ContactId, folder_id=folderId, q=query)
```

## Get Contact Folders
```python
mailbox.ContactFolders(Contact_id=ContactId, q=query)
```


