
# About

Emacs package that provides abilities to open Microsoft Exchange emails from org-mode captured links and adds events for today's and tomorrow's upcoming events to org-agenda.

# Outlook setup

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

When it's executed it copies to system's clipboard link to the currently selected message in format org-mode understands:

```
[[outlook:<internet-message-id>][EMAIL: (Subject) (Sender)]]
```

I've added a button for the macro to Outlook's ribbon and automatically it got hotkey (in my case Alt-3).

# Emacs setup

There are several variables to setup before using echange.el

```elisp
(require 'echange)

;; path of the file to save calendar data to
(setq echange/calendar-file "d:/org/calendar.org")

;; base URL of the exchange server OWA (for opening email messages in the browser)
(setq echange/exchange-base-url "https://exchange.server.com/owa")

;; path to the echange.bat file wich starts Clojure http server to access Exchange via EWS
(setq echange/server-path "./echange.bat")

;; Clojure http server port
(setq echange/server-port "5000")
```

# How it works

The link captured with VBA macro above contains internet message id of the email message. This way message is uniquely identitied and this id doesn't change when you move email to different folder in Exchange (in contrast to EntryID which is folder dependent).

When user clicks the link - internet message id is extracted from the link, passed to running http server written in Clojure (if the server is not started yet it will be started automatically) and server looks up the message by id in EWS. After it is found - it gets message's current EntryId and returns it.

After that having proper EntryId of the message we can launch Outlook with outlook:<EntryId> parameter so it opens up the message. In case when Outlook is not installed - one can open email messages in browser.

# My workflow

In Outlook select message in the list, press Alt-3, go to Emacs and use org capture capabilities to add TODO item with email link (or just add link in existing task). When you press on the link you'll be asked for credentials (only for the first time) and in couple of seconds you'll have your email message opened in Outlook.

There is also an option to open email message in browser instead of Outlook (see echange/exchange-base-url option).
