<%
'**********
'	class		: A Extensive function Library
'	File Name	: ext.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-3
'**********
'**********
'	示例
'**********

'**********
' 函数名: echo
' 参  数: str as a output string
' 作  用: Print the value of a variable
'**********
Sub echo(ByVal str)
    Response.Write str
End Sub

'**********
' 函数名: isset
' 参  数: Obj as a object
' 作  用: isNothing — Check if the object is nothing or null or undefined
'**********
Function isset(Obj)
    isset = true
    If IsEmpty(Obj) Then
        isset = false
        Exit Function
    End If
    If IsNull(Obj) Then
        isset = false
        Exit Function
    End If
	If IsObject(Obj) Then
		If Obj Is Nothing Then
			isset = false
			Exit Function
		End If
	Else
		If Obj = "" Then
			isset = false
			Exit Function
		End If
	End If
End Function

'**********
'函数名：Confirm
'作  用：确认窗口
'**********	
Function clickConfirm(msgstr)
	clickConfirm = " onclick=""javascript:return confirm('"&msgstr&"')"" "
End Function

'**********
'函数名：locationHref
'作  用：页面跳转
'**********
Sub locationHref(url)
	Response.Redirect url
	die("")
End Sub

'**********
'函数名：Referer
'作  用：返回上页
'**********
Sub locationReferer()
	locationHref(Request.ServerVariables("HTTP_REFERER"))
	die("")
End Sub

'**********
'函数名：AlertRedirect
'作  用：消息框
'**********
Sub alertRedirect(msgstr,url)
	die "<script type=""text/javascript"">"&vbCr& _
			"alert("""&msgstr&""");"&vbCr& _
			"location.replace("""&url&""");"&vbCr& _
			"</script>"
End Sub

'**********
'函数名：AlertBack
'参  数：msgstr	-- 弹出信息
'作  用：消息框
'**********
Sub alertBack(msgstr)
	die "<script type=""text/javascript"">alert(""" + msgstr + """);history.back(-1);</script>"
End Sub

'********** 
' 函数名: die
' Param: str as a output string
' 作用: Print the value of a variable and exit the procedure
'********** 
Sub die(str)
	echo(str)
	Response.End()
End Sub

'********** 
' 类名: ReAopResult
'********** 
Class ReAopResult
	Public Code
	Public Message
	Public AttachObject
End Class
%>