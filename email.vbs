strSMTPFrom = "joshuanmoses@gmail.com"
strSMTPTo = "joshuanmoses@gmail.com"
strSMTPRelay = "mail.google.com"
strTextBody = "Body of your email"
strSubject = "Test email"
'strAttachment = "full UNC path of file"


Set oMessage = CreateObject("CDO.Message")
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
oMessage.Configuration.Fields.Update

oMessage.Subject = strSubject
oMessage.From = strSMTPFrom
oMessage.To = strSMTPTo
oMessage.TextBody = strTextBody
'oMessage.AddAttachment strAttachment


oMessage.Send



'This script automatically sends messages.
