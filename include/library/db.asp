<%
'**********
'	class		: A database class
'	File Name	: db.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-4
'**********


'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Db
    Private ConnStr, SqlLocalPath
    Public ServerIp		'数据库连接主机名
	Public ConnectionType	'数据库连接类型 -- 1.ACCESS 2.MSSQL
	Public Database		'数据库名 
	Public Username			'用户名
	Public Password			'密码
    Public Conn, rs

    Private Sub Class_Initialize()
		ServerIp			= "(local)"
		ConnectionType		= "MSSQL"
		Username			= "sa"
		Password			= ""
    End Sub
	
    Private Sub Class_Terminate()
        Me.Close()
		Me.CloseRs(null)
    End Sub
	
    Sub Open()
		On Error Resume Next
		If Me.ConnectionType = "ACCESS" Then
			SqlLocalPath = Replace(Request.ServerVariables("PATH_TRANSLATED"),Replace(Request.ServerVariables("PATH_INFO"),"/","\"),"")
			ConnStr = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & SqlLocalPath & Database
		ElseIf Me.ConnectionType = "MSSQL" Then
			ConnStr = "Provider=SQLOLEDB.2;DATA SOURCE=" & Me.ServerIp & ";UID="&Username&";PWD="&Password&";Database="&Database&""
		End If
		Set Conn = CreateObject("ADODB.Connection")
		Conn.Open ConnStr
		If Err Then
			Err.Clear
			Set Conn = Nothing
			Response.Write "数据库连接出错，请检查连接字串。"
			Response.End()
		End If
		On Error GoTo 0
    End Sub
	
	Sub Close()
        On Error Resume Next
		Me.Conn.Close
		Set Me.Conn = Nothing
		Err.Clear
		On Error Goto 0
	End Sub
	
	Sub CloseRs(OutRs)
		If IsObject(OutRs) Then
			OutRs.close
			Set OutRs = Nothing
		Else
			If IsObject(rs) Then
				rs.close
				Set rs = Nothing
			End If
		End If
	End Sub

    Sub setRs( OutRs, strsql, CursorAndLockType)
		Set OutRs = Server.Createobject("ADODB.Recordset")
        OutRs.Open strsql, Conn, 1, CursorAndLockType
    End Sub

    Sub Exec(OutRs, sql)
        If InStr(UCase(sql), UCase("select"))>0 Then
            Set OutRs = Conn.Execute(sql)
        Else
            Call Conn.Execute(sql)
            OutRs = 1
        End If
    End Sub

	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程并返回记录集
	'--------------------------------------------------------------------------
	Function ExecuteRecordSet(commandName , ByVal params, ByVal conn)
		Set ExecuteRecordSet = ExecuteSqlCommand(commandName , params , 2,conn)
	End Function
	
	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程并返回记录集第一行第一列
	'--------------------------------------------------------------------------
	Function ExecuteScalar(commandName , ByVal params, ByVal conn)
		Set rs = ExecuteSqlCommand(commandName , params , 2,conn)
		If Not rs.EOF And Not rs.BOF Then
			ExecuteScalar = rs(0).Value
		Else
			ExecuteScalar = NULL
		End If
		rs.Close
		Set rs = Nothing
	End Function
	
	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程并返回一个值
	'--------------------------------------------------------------------------
	Function ExecuteReturnValue(commandName , ByVal params, ByVal conn)
		ExecuteReturnValue = ExecuteSqlCommand(commandName , params , 1, conn)
	End Function
	
	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程不返回任何内容
	'--------------------------------------------------------------------------
	Function ExecuteNonQuery(commandName , ByVal params, ByVal conn)
		Call ExecuteSqlCommand(commandName , params , 0, conn)
	End Function
	
	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程并返回记录集和一个值
	'--------------------------------------------------------------------------
	Function ExecuteRecordsetAndValue(commandName , ByVal params, ByVal conn)
		ExecuteRecordsetAndValue = ExecuteSqlCommand(commandName , params , 3, conn)
	End Function
	
	'--------------------------------------------------------------------------
	' 功能	：	执行存储过程
	'--------------------------------------------------------------------------
	' commandName		存储过程名称
	' params			参数集合，必须使用 Scripting.Dictionary 对象定义
	' returnMode		返回模式
	'					0	不返回任何参数或对象
	'					1	执行后得到返回值
	'					2	执行后得到记录集
	'					3	执行后得到返回值和记录集
	'--------------------------------------------------------------------------
	Private Function ExecuteSqlCommand(commandName , ByVal params , returnMode, ByVal conn)
		Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
		Dim iName : iName = ""
		Dim RSReturn : Set RSReturn = Nothing
		Dim ReturnValue : ReturnValue = ""
					
		cmd.ActiveConnection = conn
		cmd.CommandText = commandName
		cmd.CommandType = 4
		cmd.NamedParameters = True
	
		If returnMode = 1 Or returnMode = 3 Then
			cmd.Parameters.Append cmd.CreateParameter("@ReturnValue" , 2 , 4)
		End If
	
		If Not IsNull(params) Then
			For Each iName in params
				cmd.Parameters.Append cmd.CreateParameter(iName , 200 , 1 , 8000 , params(iName))
			Next
		End If
	
		Select Case returnMode
			' 执行后得到返回值
			Case 1
				Call cmd.Execute(, , 128)
				ExecuteSqlCommand = cmd("@ReturnValue")
				' 执行后得到记录集
			Case 2
				Set ExecuteSqlCommand = cmd.Execute()
			Case 3
				Set RSReturn = cmd.Execute()
				ExecuteSqlCommand = Array(RSReturn, cmd("@ReturnValue"))
				' 默认方式，不返回任何参数或对象
			Case Else
				Call cmd.Execute(, , 128)
		End Select
	
		Set cmd = Nothing
	End Function
End Class

%>