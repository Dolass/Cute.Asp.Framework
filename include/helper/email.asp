<%
'**********
'	class		: Md5 
'	File Name	: mg5.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-5-20
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********

Class Class_Email
	Private s_chs, s_svr, s_user, s_pwd
	
	'设置字符集
	Public Property Let Charset(str)
		s_chs = str
	End Property

	'设置字符集
	Private Property Get Charset()
		If s_chs = "" Then s_chs = "gb2312"
		s_chs = str
	End Property

	Public Property Let MailServer(str)
		s_svr = str
	End Property
	
	Private Property Get MailServer()
		If s_svr = "" Then s_svr = "12.0.0.1"
		MailServer = s_svr
	End Property

	'设置发送用户名
	Public Property Let MailUserName(str)
		s_chs = str
	End Property
	
	Private Property Get MailUserName()
		MailUserName = s_user
	End Property

	'设置发送用户的密码
	Public Property Let MailPassword(str)
		s_pwd = str
	End Property
	
	Private Property Get MailPassword()
		MailPassword = s_pwd
	End Property

	'**********
    '函数名：SendMail
    '作  用：用jMail组件发送邮件
    '参  数：MailtoAddress ----收信人地址
    '        MailtoName    -----收信人姓名
    '        Subject       -----主题
    '        TemplateFile  -----模板文件路径
    '        Params		   -----替换参数
    '        FromName      -----发信人姓名
    '        MailFrom      -----发信人地址
    '        Priority      -----信件优先级（1为加急，3为普通，5为低级）
	'返回值：Array(是否成功,提示信息)
    '**********
    Function SendMail(MailtoAddress, MailtoName, Subject, TemplateFile, Params, FromName, FromMail, Priority)
        If IsObjInstalled("jMail.Message") Then
            On Error Resume Next
        Else
			SendMail = Array(False,"未找到 JMail 组件")
            Exit Function
        End If
		If Priority = "" Or IsNull(Priority) Then Priority = 3

		Dim jMail : Set jMail = Server.CreateObject("jMail.Message")
        jMail.Silent = False
        jMail.Charset = System_Charset					'邮件编码
        jMail.ContentType = "text/html"					'邮件正文格式
        jMail.From = FromMail							'发信人Email
        jMail.FromName = FromName						'发信人姓名
		jMail.ReplyTo = FromMail						
		jMail.AppendBodyFromFile TemplateFile			'内容模板

		Dim iParamName
		If Not IsNull(Params) Then
			For Each iParamName In Params
				jMail.Body = Replace(jMail.Body, "{" & iParamName & "}", Params(iParamName))
			Next
		End If
        
		jMail.ClearRecipients
        jMail.AddRecipient MailtoAddress, MailtoName	'收信人
        jMail.Subject = Subject							'主题
        jMail.Priority = Priority						'邮件等级，1为加急，3为普通，5为低级

		'如果服务器需要SMTP身份验证则还需指定以下参数
        jMail.MailDomain = MailDomain					'域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
        jMail.MailServerUserName = MailUserName			'登录用户名
        jMail.MailServerPassWord = MailPassword			'登录密码

        jMail.Send MailServer
        SendMail = jMail.ErrorMessage
        jMail.Close
        Set jMail = Nothing
		If Err Then
			SendMail = Array(False,Err.Description)
			Err.Clear
			Exit Function
		End If
		SendMail = Array(True,SendMail)
        On Error Goto 0
    End Function

    '**********
    '函数名：IsObjInstalled
    '作  用：检查组件是否已经安装
    '参  数：strClassString ----组件名
    '返回值：True  ----已经安装
    '       False ----没有安装
    '**********
    Private Function IsObjInstalled(strClassString)
        IsObjInstalled = False
        On Error Resume Next
        Dim xTestObj
        Set xTestObj = Server.CreateObject(strClassString)
        If Err.Number = 0 Then IsObjInstalled = True
        Set xTestObj = Nothing
        On Error Goto 0
    End Function

End Class

%>
