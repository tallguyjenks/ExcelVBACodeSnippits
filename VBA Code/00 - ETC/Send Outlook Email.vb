Sub SendEmailFromOutlook(body As String, subject As String, toEmails As String, ccEmails As String, bccEmails As String)
    Dim outApp As Object
    Dim outMail As Object
    Set outApp = CreateObject(""Outlook.Application"")
    Set outMail = outApp.CreateItem(0)
 
    With outMail
        .to = toEmails
        .CC = ccEmails
        .BCC = bccEmails
        .subject = subject
        .HTMLBody = body
        .Send 'Send the email
    End With
 
    Set outMail = Nothing
    Set outApp = Nothing
End Sub


Sub SomeMacro()

'insert macro to run here

'(body, subject, to, CC, Bcc)
    Call SendEmailFromOutlook("""", """", """", """", """")
End Sub
