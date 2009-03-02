<%
'**********
'	class		: some function Library
'	File Name	: function.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-3
'**********

'**********
'函数名：showErr
'参  数：message	-- 错误信息
'作  用：显示错误信息
'**********
Sub showErr(message)
    die "<html><head><title>Exception page</title><meta http-equiv=""Content-Type"" content=""text/html; charset=gbk"" /><style type=""text/css""><!--" & vbNewLine & "* { margin:0; padding:0 }" & vbNewLine & "body { background:#333; color:#0f0; font:14px/1.6em ""宋体"", Verdana, Arial, Helvetica, sans-serif; }" & vbNewLine & "dl { margin:20px 40px; padding:20px; border:3px solid #f63; }" & vbNewLine & "dt { margin:0 0 0.8em 0; font-weight:bold; font-size:1.6em; }" & vbNewLine & "dd { margin-left:2em; margin-top:0.2em; }" & vbNewLine & "--></style></head><body><div id=""container""><dl><dt>Description:</dt><dd><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Position:</span> " & message & "</dd></dl></div></body></html>"
End Sub

'**********
'函数名：showException
'作  用：显示异常信息
'**********
Sub showException()
    echo "<p><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Error:</span> " & Err.Number & " " & Err.Description & "</p>"
    Err.Clear
    die("")
End Sub

'**********
' 函数名: checkPostSource
' 作  用: 检验来源地址
'**********
Function checkPostSource()
	Dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(server_v1,8,Len(server_v2))=server_v2 Then
		checkPostSource=True
	Else
		checkPostSource=False
	End If
End Function

'**********
'函数名：GetSystem
'作  用：获取客户端操作系统版本
'返回值：操作系统版本名称
'**********
Function GetSystem()
	Dim System
	System = Request.ServerVariables("HTTP_USER_AGENT")
	If Instr(System,"Windows NT 5.2") Then
		System = "Win2003"
	ElseIf Instr(System,"Windows NT 5.0") Then
		System="Win2000"
	ElseIf Instr(System,"Windows NT 5.1") Then
		System = "WinXP"
	ElseIf Instr(System,"Windows NT") Then
		System = "WinNT"
	ElseIf Instr(System,"Windows 9") Then
		System = "Win9x"
	Elseif Instr(System,"unix") Or InStr(System,"linux") Or InStr(System,"SunOS") Or InStr(System,"BSD") Then
		System = "Unix"
	ElseIf Instr(System,"Mac") Then
		System = "Mac"
	Else
		System = "Other"
	End If
	GetSystem = System
End Function

'**********
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：obj ----组件名
'返回值：True  ----已经安装
'		 False ----没有安装
'**********
Function IsObjInstalled(obj)
	On Error Resume Next
	IsObjInstalled = False
	Dim xTestObj
	Set xTestObj = Server.CreateObject(obj)
	If Err.Number = 0 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err.Clear
	On Error Goto 0
End Function

'**********
' 函数名: noBuffer
' 作  用: no buffer
'**********
Sub noBuffer()
	Response.Buffer = True
	Response.Expires = 0
	Response.AddHeader "Expires",-1
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","no-cache,must-revalidate"
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
End Sub

'**********
' 函数名: rand
' 参  数: str as the input string
' 作  用: Generate a random integer
'**********
Function rand(min, max)
    Randomize
    rand = Int((max - min + 1) * Rnd + min)
End Function

'**********
' 函数名: rq
' 参  数: Requester as the request type
' 参  数: Name as the request name
' 参  数: iType as check type
' 参  数: Default as the Default string
' 作  用: safe filter
'**********
Function rq(Requester,Name,iType,Default)
	Dim tmp
	Select Case Requester
	Case 0
		tmp = Name
	Case 1
		tmp = Trim(Request(Name))
		tmp = htmlEncode(tmp)
	Case 2
		tmp = Trim(Request.QueryString(Name))
		tmp = htmlEncode(tmp)
	Case 3
		tmp = Trim(Form(Name))
		tmp = htmlEncode(tmp)
	Case 4
		tmp = Request.Cookies(Name)
	End Select
	If tmp = "" Then tmp = Default
	Select Case iType
	Case 0
		If IsNumeric(tmp) = False Then
			tmp = Default
		Else
			tmp = CSng(tmp)
		End If
	Case 1
		tmp = safe(encodeJP(tmp))
	Case 2
		If Not IsDate(tmp) Or Len(tmp) <= 0 Then 
			tmp = CDate(Default)
		Else
			tmp = CDate(tmp)
		End If
	End Select
	rq = tmp
End function

'**********
'函数名：Form
'参  数：element ---- 控件名
'作  用：获取Form控件数据
'**********
Function Form(element)
    On Error Resume Next
    If InStr(LCase(Request.ServerVariables("Content_Type")), "multipart/form-data") Then	'multipart/form-data
        If IsObject(Tpub.Upload) = False Then
    		If Err Then
				die("no include upload class file")
			End If
        End If
		If Tpub.Upload.Error <> 0 Then
			Tpub.Upload.Open
		End If
        Form = Tpub.Upload.Form(element)
    Else
        Form = Request.Form(element)
    End If
    On Error GoTo 0
End Function

'**********
'函数名：sqlFilter
'作  用：过滤Sql关键字
'**********
Function safe(str)
	If str = "" Then Exit Function
	str = Replace(str, "'", "''")
	str = Replace(str,"--","&#45;&#45;")
	safe = str
End Function

'**********
' 函数名: currentURL
' 作  用: 返回当前地址
'**********
Function currentURL()
	Dim port : port = LCase(Request.ServerVariables("Server_Port"))
    Dim page : page = LCase(Request.ServerVariables("Script_Name"))
    Dim query : query = LCase(Request.QueryString())
	Dim url
	If CStr(port) = "80" Then 
		url = page 
	Else
		url = ":" & port & page
	End If
	If query <> "" Then
		currentURL = "http://" & Request.ServerVariables("server_name") & url & "?" & query
	Else
		currentURL = "http://" & Request.ServerVariables("server_name") & url
	End If
End Function

'**********
' 函数名: refererURL
' 作  用: 返回来源地址
'**********
Function refererURL()
	refererURL = Request.ServerVariables("HTTP_REFERER")
End Function

'**********
' 函数名: getIP
' 作  用: 获取客户端IP
'**********
Function getIP()
	Dim Ip,Tmp
	Dim i,IsErr
	Dim ForTotal
	IsErr=False
	Ip=Request.ServerVariables("REMOTE_ADDR")
	If Len(Ip)<=0 Then Ip=Request.ServerVariables("HTTP_X_ForWARDED_For")
	If Len(Ip)>15 Then 
		IsErr=True
	Else
		Tmp=Split(Ip,".")
		If Ubound(Tmp)=3 Then 
			ForTotal = Ubound(Tmp)
			For i=0 To ForTotal
				If Len(Tmp(i))>3 Then IsErr=True
			Next
		Else
			IsErr=True
		End If
	End If
	If IsErr Then 
		getIP="1.1.1.1"
	Else
		getIP=Ip
	End If
End Function

'**********
' 函数名: basename
' 作  用: 获取当前访问文件名
'**********
Function getSelfName()
    getSelfName = Request.ServerVariables("PATH_TRANSLATED")
    getSelfName = Mid(getSelfName, InstrRev(getSelfName, "\") + 1, Len(getSelfName))
End Function

'**********
'函数名：ReturnObj
'作  用：返回一个对象
'返回值：对象包含三个参数(Code,Message,AttachObject)
'**********
Function returnObj()
	On Error Resume Next
	TypeName(New AopResult)
	If Err Then
		Set returnObj = New ReAopResult		'重定义的AopResult
		Err.Clear
	Else
		Set returnObj = New AopResult
	End If
	On Error Goto 0
End Function

'**********
' 函数名: encodeJP
' 参  数: str as the input string
' 作  用: 编码日文
'**********
Function encodeJP(ByVal str)
	If str="" Then Exit Function
	Dim c1 : c1 = Array("ガ","ギ","グ","ア","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","パ","ビ","ピ","ブ","プ","ベ","ペ","ボ","ポ","ヴ")
	Dim c2 : c2 = Array("460","462","463","450","466","468","470","472","474","476","478","480","482","485","487","489","496","497","499","500","502","503","505","506","508","509","532")
	Dim i
	For i=0 to 26
		str=Replace(str,c1(i),"&#12"&c2(i)&";")
	Next
	encodeJP = str
End Function

'**********
' 函数名: htmlEncode
' 参  数: str as the input string
' 作  用: filter html code
'**********
Function htmlEncode(ByVal Str)
	If Trim(Str) = "" Or IsNull(Str) Then
		htmlEncode = ""
	Else
		str = Replace(str, "  ", "&nbsp; ")
		str = Replace(str, """", "&quot;")
		str = Replace(str, ">", "&gt;")
		str = Replace(str, "<", "&lt;")
		htmlEncode = Str
	End If
End Function

'**********
' 函数名: htmlDecode
' 参  数: str as the input string
' 作  用: Decode the html tag
'**********
Function htmlDecode(ByVal str)
	If Not IsNull(str) And str <> "" Then
		str = Replace(str, "&nbsp;", " ")
		str = Replace(str, "&quot;", """")
		str = Replace(str, "&gt;", ">")
		str = Replace(str, "&lt;", "<")
		htmlDecode = str
	End If
End Function

'**********
' 函数名: URLDecode
' 作  用: URLDecode — URL decode
'**********
Function URLDecode(ByVal vstrin)
	Dim i, strreturn, strSpecial, intasc, thischr
	strSpecial = "!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%"
	strreturn = ""
	For i = 1 To Len(vstrin)
		thischr = Mid(vstrin, i, 1)
		If thischr = "%" Then
			intasc = Eval("&h" + Mid(vstrin, i + 1, 2))
			If InStr(strSpecial, Chr(intasc))>0 Then
				strreturn = strreturn & Chr(intasc)
				i = i + 2
			Else
				intasc = Eval("&h" + Mid(vstrin, i + 1, 2) + Mid(vstrin, i + 4, 2))
				strreturn = strreturn & Chr(intasc)
				i = i + 5
			End If
		Else
			If thischr = "+" Then
				strreturn = strreturn & " "
			Else
				strreturn = strreturn & thischr
			End If
		End If
	Next
	URLDecode = strreturn
End Function

'**********
' 函数名: randStr
' 作用: Generate a specific length random string
'**********
Function randStr(intLength)
	Dim strSeed,seedLength,i
	strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
	seedLength = len(strSeed)
	For i=1 to intLength
		Randomize
		randStr = randStr & Mid(strSeed,Round((Rnd*(seedLength-1))+1),1)
	Next
End Function

'**********
' 函数名: SetQueryString
' 作用: 重置参数
'**********
Function SetQueryString(ByVal sQuery, ByVal Name,ByVal Value)
	Dim Obj
	If Len(sQuery) > 0 Then
		If InStr(1,sQuery,Name&"=",1) = 0 Then
			If InStr(sQuery,"=") > 0 Then
				sQuery = sQuery & "&" & Name & "=" & Value
			Else
				sQuery = sQuery & Name & "=" & Value
			End If
		Else
			Set Obj = New Regexp
			Obj.IgnoreCase = False
			Obj.Global = True
			Obj.Pattern = "(" & Name & "=)[^&]+"
			sQuery = Obj.Replace(sQuery,"$1" & Value)
			Set Obj = Nothing
		End If
	Else
		sQuery = sQuery & Name & "=" & Value
	End If
	SetQueryString = sQuery
End Function
%>
