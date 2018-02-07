About:

Emacs package that provides abilities to open Microsoft Exchange emails from org-mode captured links and adds events for today's and tomorrow's upcoming events to org-agenda.

Outlook setup:

In Microsoft Outlook create the following macro:

```
Sub CopyLinkToClipboard()
   Dim objMail As Outlook.MailItem
   Dim doClipboard As New DataObject

   'One and ONLY one message muse be selected
   If Application.ActiveExplorer.Selection.Count <> 1 Then
       MsgBox ("Select one and ONLY one message.")
       Exit Sub
   End If

   Set objMail = Application.ActiveExplorer.Selection.Item(1)
   doClipboard.SetText "[[outlook:" + objMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") + "][EMAIL: " + objMail.Subject + " (" + objMail.SenderName + ")]]"
   doClipboard.PutInClipboard
End Sub
```

When it's executed it copies to clipboard link to the message in format org-mode understands:

```
[[outlook:<internet-message-id>][EMAIL: (Subject) (Sender)]]
```
