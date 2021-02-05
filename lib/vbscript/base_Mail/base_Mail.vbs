Option Explicit

Const SEND_USING_PICKUP = 1    	' Send message using local SMTP service pickup directory
Const SEND_USING_PORT 	= 2  	' Send the message using SMTP over TCP/IP networking

Const AUTH_ANONYMOUS 	= 0  	' No authentication
Const AUTH_BASIC 	= 1  	' BASIC clear text authentication
Const AUTH_NTLM		= 2   	' NTLM, Microsoft proprietary authentication

Class base_Mail
	Private p_Mail

	Private Sub Class_Initialize()
		Set p_Mail = CreateObject("CDO.Message")
	End Sub


	' Properties


	Public Property Get Attachments()
		Set Attachments = p_Mail.Attachments 
	End Property

	Public Property Get Authentication()
		Authentication = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value
	End Property

	Public Property Let Authentication(intAuthentication)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate").Value = intAuthentication 
	End Property

	Public Property Get AutoGenerateTextBody() 
		AutoGenerateTextBody = p_Mail.AutoGenerateTextBody
	End Property

	Public Property Let AutoGenerateTextBody(blnAutoGenerateTextBody)
		p_Mail.AutoGenerateTextBody = blnAutoGenerateTextBody
	End Property

	Public Property Get Bcc() 
		Bcc = p_Mail.BCC
	End Property

	Public Property Let Bcc(strBcc)
		p_Mail.BCC = strBcc
	End Property

	Public Property Get BodyPart()
		Set BodyPart = p_Mail.BodyPart
	End Property

	Public Property Get Cc() 
		Cc = p_Mail.CC
	End Property

	Public Property Let Cc(strCc) 
		p_Mail.CC = strCc
	End Property

	Public Property Get Configuration() 
		Set Configuration = p_Mail.Configuration
	End Property

	Public Property Set Configuration(objConfiguration)
		Set p_Mail.Configuration = objConfiguration
	End Property

	Public Property Get DataSource() 
		Set DataSource = p_Mail.DataSource
	End Property

	Public Property Get DsnOptions() 
		Set DsnOptions = p_Mail.DSNOptions
	End Property

	Public Property Set DsnOptions(objCdoDsnOptions)
		Set p_Mail.DSNOptions = objCdoDsnOptions
	End Property

	Public Property Get EnvelopeFields() 
		Set EnvelopeFields = p_Mail.EnvelopeFields
	End Property

	Public Property Get Fields() 
		Set Fields = p_Mail.Fields
	End Property

	Public Property Get FollowUpTo()
		FollowUpTo = p_Mail.FollowUpTo
	End Property

	Public Property Let FollowUpTo(strFollowUpTo)
		p_Mail.FollowUpTo = strFollowUpTo
	End Property

	Public Property Get From()
		From = p_Mail.From
	End Property

	Public Property Let From(strFrom)
		p_Mail.From = strFrom
	End Property

	Public Property Get HtmlBody()
		HtmlBody = p_Mail.HTMLBody
	End Property

	Public Property Let HtmlBody(strHtmlBody)
		p_Mail.HTMLBody = strHtmlBody
	End Property

	Public Property Get HtmlBodyPart() 
		Set HtmlBodyPart = p_Mail.HTMLBodyPart
	End Property

	Public Property Get Keywords()
		Keywords = p_Mail.Keywords
	End Property

	Public Property Let Keywords(strKeywords)
		p_Mail.Keywords = strKeywords
	End Property

	Public Property Get MdnRequested()
		MdnRequested = p_Mail.MDNRequested
	End Property

	Public Property Let MdnRequested(blnMdnRequested)
		p_Mail.MDNRequested = blnMdnRequested
	End Property

	Public Property Get MimeFormatted()
		MimeFormatted = p_Mail.MimeFormatted
	End Property

	Public Property Let MimeFormatted(blnMimeFormatted)
		p_Mail.MimeFormatted = blnMimeFormatted
	End Property

	Public Property Get Newsgroups()
		Newsgroups = p_Mail.Newsgroups
	End Property

	Public Property Let Newsgroups(strNewsgroups)
		p_Mail.Newsgroups = strNewsgroups
	End Property

	Public Property Get Organization()
		Organization = p_Mail.Organization
	End Property

	Public Property Let Organization(strOrganization)
		p_Mail.Organization = strOrganization
	End Property

	Public Property Get Password()
		UserName = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value
	End Property

	Public Property Let Password(strPassword)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword").Value = strPassword
	End Property

	Public Property Get ReceivedTime()
		ReceivedTime = p_Mail.ReceivedTime
	End Property

	Public Property Get Recipient()
		Recipient = p_Mail.To
	End Property

	Public Property Let Recipient(strRecipient)
		p_Mail.To = strRecipient
	End Property

	Public Property Get ReplyTo()
		ReplyTo = p_Mail.ReplyTo
	End Property

	Public Property Let ReplyTo(strReplyTo)
		p_Mail.ReplyTo = strReplyTo
	End Property

	Public Property Get Sender()
		Sender = p_Mail.Sender
	End Property

	Public Property Let Sender(strSender)
		p_Mail.Sender = strSender
	End Property

	' SendUsing

	Public Property Get SentOn() 
		SentOn = p_Mail.SentOn
	End Property

	Public Property Get SmtpServer()
		SmtpServer = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value
	End Property

	Public Property Let SmtpServer(strSmtpServer)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = strSmtpServer
	End Property

	Public Property Get SmtpServerPort()
		SmtpServerPort = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value
	End Property

	Public Property Let SmtpServerPort(intServerPort)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = intServerPort
	End Property

	Public Property Get Ssl()
		Ssl = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl").Value
	End Property

	Public Property Let Ssl(blnSsl)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl").Value = blnSsl
	End Property

	Public Property Get Subject()
		Subject = p_Mail.Subject
	End Property

	Public Property Let Subject(strSubject)
		p_Mail.Subject = strSubject
	End Property

	Public Property Get TextBody()
		TextBody = p_Mail.TextBody
	End Property

	Public Property Let TextBody(strTextBody)
		p_Mail.TextBody = strTextBody
	End Property

	Public Property Get TextBodyPart()
		Set TextBodyPart = p_Mail.TextBodyPart
	End Property
	
	Public Property Get Timeout()
		TextBody = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout").Value
	End Property

	Public Property Let Timeout(intTimeout)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout").Value = intTimeout
	End Property

	Public Property Get Username()
		UserName = p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername").Value
	End Property

	Public Property Let Username(strUsername)
		p_Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername").Value = strUsername
	End Property	


	' Methods


	Public Function AddAttachment(strUrl) ' Optional params: [UserName As String], [Password As String]) As IBodyPart
		Set AddAttachment = p_Mail.AddAttachment(strUrl)
	End Function

	Public Function AddRelatedBodyPart(strUrl, strReference, objReferenceType) ' Optional params: [UserName As String], [Password As String]) As IBodyPart
		Set AddRelatedBodyPart = p_Mail.AddRelatedBodyPart(strUrl, strReference, objReferenceType)
	End Function
 
	Public Sub CreateMHTMLBody(strUrl) ' Optional params: [Flags As CdoMHTMLFlags = cdoSuppressNone], [UserName As String], [Password As String])
		p_Mail.CreateMHTMLBody strUrl
	End Sub

	Public Function Forward()
		Set Forward = p_Mail.Forward()
	End Function

	Public Function GetInterface(strInterface) 
		Set Forward = p_Mail.GetInterface(strInterface)
	End Function

	Public Function GetStream() 
		Set GetStream = p_Mail.GetStream()
	End Function

	Public Sub Post() 
		p_Mail.Post
	End Sub

	Public Function PostReply()
		Set PostReply = p_Mail.PostReply()
	End Function

	Public Function Reply() 
		Set Reply = p_Mail.Reply()
	End Function

	Public Function ReplyAll() 
		Set ReplyAll = p_Mail.ReplyAll()
	End Function

	Public Sub Send() 
		p_Mail.Send
	End Sub

	Private Sub Class_Terminate()
		Set p_Mail = Nothing
	End Sub
End Class

If WScript.ScriptName = "base_Mail.vbs" Then

End If
