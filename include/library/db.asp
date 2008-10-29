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
	
    '**********
    ' 方法名: Open
    ' 作  用: 打开数据库链接
    '**********
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

    '**********
    ' 方法名: Close
    ' 作  用: 关闭并释放数据库链接
    '**********
	Sub Close()
        On Error Resume Next
		Me.Conn.Close
		Set Me.Conn = Nothing
		Err.Clear
		On Error Goto 0
	End Sub
	
    '**********
    ' 方法名: CloseRs
    ' 参  数: OutRs as recordset object
    ' 作  用: 关闭并释放指定对象
    '**********
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

    '**********
    ' 方法名: Exec
    ' 参  数: OutRs as recordset object
    ' 参  数: sql as sql string
    ' 作  用: 执行sql语句，若是查询语句则将结果返回给OutRs
    '**********
    Sub Exec(OutRs, sql)
        If InStr(UCase(sql), UCase("select"))>0 Then
            Set OutRs = Conn.Execute(sql)
        Else
            Call Conn.Execute(sql)
            OutRs = 1
        End If
    End Sub

    '**********
    ' 方法名: GetRecordObject
    ' 参  数: source as sql string or recordset object
    ' 作  用: 将一条记录转为一个字典对象
    '**********
	Function GetRecordObject(source)
		Dim Params, i
		If TypeName(rs) <> "Recordset" Then
			Tpub.db.Exec rs,source
		End If
		If rs.eof Then 
			Set GetEntryObject = Nothing
			Exit Function
		End If
		Set Params = Server.CreateObject("Scripting.Dictionary")
		Params.CompareMode = 1
		For i=0 To rs.fields.count-1
			If IsNumeric(rs(i)) Then
				If TypeName(Not rs(i)) = "Boolean" Then
					Params.Add rs.Fields(i).Name,Not Not rs(i)
				Else
					Params.Add rs.Fields(i).Name,rs(i)+0
				End If
			Else
				Params.Add rs.Fields(i).Name,rs(i)&""
			End if
		Next
		Tpub.db.closeRs(rs)
		Set GetRecordObject = Params
		Set Params = Nothing
	End Function

    '**********
    ' 方法名: Update
    ' 参  数: Table as Data Table
    ' 参  数: Params as Dictionary
    ' 参  数: Where as 条件语句
    ' 作  用: 更新记录
    '**********
	Function Update(Table,Params,Where)
		Dim sqlCmd, parameteres, oParams
		Dim iName
		sqlCmd="Update "&Table&" set "
		parameteres = " "
		Set oParams = Server.CreateObject("Scripting.Dictionary")
		If Not IsNull(params) Then
			For Each iName in params
				sqlCmd = sqlCmd & iName & "=@" & iName & ","
				parameteres = parameteres & "@" & iName & " varchar(8000)" & ","
			Next
		End If
		sqlCmd = Left(sqlCmd,Len(sqlCmd)-1)
		If Trim(Where) <> "" Then sqlCmd=sqlCmd&" Where "&Where&""
		parameteres = Left(parameteres,Len(parameteres)-1)
		oParams.Add "@stmt",sqlCmd
		oParams.Add "@parameters",parameteres
		If Not IsNull(Params) Then
			For Each iName in Params
				oParams.Add "@"&iName,Params(iName)
			Next
		End If
		Call Me.ExecuteNonQuery("sp_executesql",oParams)
		Set Params = Nothing
		Set oParams = Nothing
	End Function

    '**********
    ' 方法名: Insert
    ' 参  数: Table as Data Table
    ' 参  数: Params as Dictionary
    ' 参  数: Where as 条件语句
    ' 作  用: 插入记录
    '**********
	Function Insert(Table,Params)
		Dim sqlCmd, sqlCmd_a, sqlCmd_b, parameteres, oParams
		Dim iName
		sqlCmd = "Insert Into "&Table&" ("
		parameteres = " "
		Set oParams = Server.CreateObject("Scripting.Dictionary")
		If Not IsNull(params) Then
			For Each iName in params
				sqlCmd_a = sqlCmd_a & iName & ","
				sqlCmd_b = sqlCmd_b & "@" & iName & ","
				parameteres = parameteres & "@" & iName & " varchar(8000)" & ","
			Next
		End If
		sqlCmd_a = Left(sqlCmd_a,Len(sqlCmd_a)-1)
		sqlCmd_b = Left(sqlCmd_b,Len(sqlCmd_b)-1)
		sqlCmd = sqlCmd & sqlCmd_a & ")  values(" & sqlCmd_b & ")"
		parameteres = Left(parameteres,Len(parameteres)-1)
		oParams.Add "@stmt",sqlCmd
		oParams.Add "@parameters",parameteres
		If Not IsNull(Params) Then
			For Each iName in Params
				oParams.Add "@"&iName,Params(iName)
			Next
		End If
		Call Me.ExecuteNonQuery("sp_executesql",oParams)
		Set Params = Nothing
		Set oParams = Nothing
	End Function

	'**********
	' 功能	：	执行存储过程并返回记录集
	'**********
	Function ExecuteRecordSet(commandName , ByVal params)
		Set ExecuteRecordSet = ExecuteSqlCommand(commandName , params , 2)
	End Function
	
	'**********
	' 功能	：	执行存储过程并返回记录集第一行第一列
	'**********
	Function ExecuteScalar(commandName , ByVal params)
		Set rs = ExecuteSqlCommand(commandName , params , 2)
		If Not rs.EOF And Not rs.BOF Then
			ExecuteScalar = rs(0).Value
		Else
			ExecuteScalar = NULL
		End If
		rs.Close
		Set rs = Nothing
	End Function
	
	'**********
	' 功能	：	执行存储过程并返回一个值
	'**********
	Function ExecuteReturnValue(commandName , ByVal params)
		ExecuteReturnValue = ExecuteSqlCommand(commandName , params , 1)
	End Function
	
	'**********
	' 功能	：	执行存储过程不返回任何内容
	'**********
	Function ExecuteNonQuery(commandName , ByVal params)
		Call ExecuteSqlCommand(commandName , params , 0)
	End Function
	
	'**********
	' 功能	：	执行存储过程并返回记录集和一个值
	'**********
	Function ExecuteRecordsetAndValue(commandName , ByVal params)
		ExecuteRecordsetAndValue = ExecuteSqlCommand(commandName , params , 3)
	End Function
	
	'**********
	' 功能	：	执行存储过程
	'**********
	' commandName		存储过程名称
	' params			参数集合，必须使用 Scripting.Dictionary 对象定义
	' returnMode		返回模式
	'					0	不返回任何参数或对象
	'					1	执行后得到返回值
	'					2	执行后得到记录集
	'					3	执行后得到返回值和记录集
	'**********
	Private Function ExecuteSqlCommand(commandName , ByVal params , returnMode)
		Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
		Dim iName : iName = ""
		Dim RSReturn : Set RSReturn = Nothing
		Dim ReturnValue : ReturnValue = ""
		
		cmd.ActiveConnection = Me.conn
		cmd.CommandText = commandName
		cmd.CommandType = 4
		cmd.NamedParameters = True
	
		If returnMode = 1 Or returnMode = 3 Then
			cmd.Parameters.Append cmd.CreateParameter("@ReturnValue" , 2 , 4)
		End If
	
		If Not IsNull(params) Then
			For Each iName in params
				If iName <> "@stmt" And iName <> "@statement" And iName <> "@parameters" Then
					cmd.Parameters.Append cmd.CreateParameter(iName , 200 , 1 , 8000 , params(iName))
				Else
					cmd.Parameters.Append cmd.CreateParameter(iName , 203 , 1 , 4000 , params(iName))
				End If
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