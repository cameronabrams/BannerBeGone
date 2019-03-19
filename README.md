# BannerBeGone
Instructions on how to remove email warning banners in Outlook @ Drexel

On or about March 11, 2019, Drexel IT began to tag all non-Drexel incoming email with a warning banner at the top of the email body that looks like this:

<div style="border:solid black 1.0pt;margin-left:2.0pt;margin-top:2.0pt;margin-right:2.0pt;margin-bottom:6.0pt">
<p style="margin:2.0pt"><b><span style="color:maroon">Caution: This message came from outside of Drexel.</span></b>
<u>Do not click links or attachments</u> unless you <em>expected</em> this email.</p>
</div>

Apart from being ugly and annoying, it is downright insulting to those of us who understand security and are aware of phishing.  To combat this annoyance, I figured out a way to strip these banners from all incoming messages using the Rules feature of Outlook.  Unfortunately, this only works for the Outlook client on my Windows laptop, but it is a start.  Basically, it amounts to running a custom script on every email received to remove the banner text if it is encountered.

1. In Outlook, Alt-F11 to open the VBA interface.  Then create in Project1->'Microsoft Outlook Objects'->ThisOutlookSession and paste in this code:

```
Option Explicit

Sub InsertHyperLink(MyMail As MailItem)

    Dim body As String

    body = MyMail.HTMLBody

    body = Replace(body, "border:solid black 1.0pt", "", 1, 1, vbTextCompare)
    body = Replace(body, "Caution: This message came from outside of Drexel.", "", 1, 1, vbTextCompare)
    body = Replace(body, "<u>Do not click links or attachments</u> unless you <em>expected</em> this email.", "", 1, 1, vbTextCompare)

    MyMail.HTMLBody = body

    MyMail.Save

End Sub
```

2. Now, you need to sign this with a certificate.  If you don't have one, you can make your own:
   1. In powershell, navigate to `C:\Program Files (x86)\Microsoft Office\Office15`
   2. Issue `SELFCERT.EXE` and create your own certificate
   3. In IE, 'Internet Options'->Content->Certificates, export this personal certificate and then imported into the trusted root certificates.
   4. Sign your VBA script with this certificate (Tools->Digital Signature)

3. Now, you need to edit your windows registry.  
   1. Invoke `regedit` from the command line
   2. Go to `HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Security`
   3. Create a new DWORD; name it `EnableUnsafeClientMailRules` and give it the value 1.

4. Finally, add the rule in Outlook.
   1. Manange Rules->New Rule->'Apply rule on messages I receive'
   2. No condition (all recevied messages will be subjected to the action)
   3. At the 'What do you and to do with the message? Select action(s):' check "run a script", and then click on the word "script" in the Step 2 window, and select your `Project1.ThisOutSession.Inse...` script.
   4. Make a little folder in your mailbox and put some existing externally-sourced emails with the annoying banner in there.   Run the script on those messages and verify the banner is removed cleanly.

