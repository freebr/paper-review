<%
	Class Messenger
    Dim m_msgType
    Property Get MsgType
    	MsgType = m_msgType
    End Property
    Property Let MsgType(value)
        postUrl = msgs(value).Attributes.getNamedItem("postUrl").Value
        account = msgs(value).Attributes.getNamedItem("account").Value
        password = msgs(value).Attributes.getNamedItem("password").Value
		End Property
		Dim postUrl,account,password
		Dim xmlConfig
    Dim msgs
    Public Sub Class_Initialize()
    	Set xmlConfig = Server.CreateObject("MSXML2.DomDocument.4.0")
			xmlConfig.preserveWhiteSpace=False
      If Not xmlConfig.Load(Server.MapPath("/config/sms.config")) Then Exit Sub
      Set msgs = xmlConfig.SelectNodes("configuration/messengers/messenger")
      MsgType = 0
    End Sub
		Public Function sendMessage(recipient, content)
      recipient = Server.URLEncode(Trim(recipient))
      content = Server.URLEncode(Trim(content))
      Dim paramTemplate
      paramTemplate = "account={0}&pswd={1}&mobile={2}&msg={3}&needstatus=true&product=&extno="
      Dim postData
      postData = Replace(Replace(Replace(Replace(paramTemplate,"{0}",account), "{1}", password), "{2}", recipient), "{3}", content)
      
      Dim req:Set req=Server.CreateObject("MSXML2.XMLHTTP")
      req.Open "POST", postUrl, false
      req.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
      req.Send postData
      If req.Status=200 Then
       	Dim rp:rp = req.ResponseText
        Dim result:result = Split(rp,vbLf)(0)
        sendMessage = Split(result,",")(1)
      Else	' 访问失败
        sendMessage = 0
      End If
    End Function
	End Class
%>