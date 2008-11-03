<%
'**********
'	class		: page
'	File Name	: page.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-6-20
'**********


'**********
'	示例
'**********

'Dim Getdata
'Tpub.page.Conn = Tpub.db.conn
'Tpub.page.Page = Trim(Request("PageID"))
'Tpub.page.Size = 10
'Tpub.page.Header_b(rs,sTable,sPK,sFields,sWhere,sSort)
'Do While Not rs.eof
'	--------
'	MoveNext
'Loop
'Tpub.page.Footer_a "",""

'**********
'	构建类
'**********

Class Class_Page
	Private i_pCount,i_rCount,i_pSize,i_pNumber,i_conn,CurrentPath
	Public PageProcedure

	Private Sub Class_Initialize()
		i_pSize = 1
		i_pNumber = 1
		i_pCount = 0
		i_rCount = 0
		Dim port : port = LCase(Request.ServerVariables("Server_Port"))
		Dim page : page = LCase(Request.ServerVariables("Script_Name"))
		Dim url
		If CStr(port) = "80" Then 
			CurrentPath = "http://" & Request.ServerVariables("server_name") & page 
		Else
			CurrentPath = "http://" & Request.ServerVariables("server_name") & ":" & port & page
		End If
		CurrentPath = CurrentPath & "?" & regExpReplace(Request.QueryString(),"&?pageid=\d+[^\&]*","",false)
		CurrentPath = regExpReplace(CurrentPath,"(^http\:\/\/.+[^&\?]+)$","$1&",false)
		CurrentPath = Replace(CurrentPath,"?&","?")
		PageProcedure = "PagingLarge"
	End Sub
	
	'**********属性设置**************************
	'Conn : 数据库链接对象
	'Size : 每页显示
	'Page : 当前页码
	'**********
	Public Property Let Conn(ByVal n)
		Set i_conn = n
	End Property
	
	Public Property Let Size(ByVal n)
		i_pSize = n
	End Property

	Public Property Let Page(ByVal n)
		If Not IsNumeric(n) Then n = 1
		If n <= 0 Then n = 1
		i_pNumber = CLng(n)
	End Property
	
	'**********获取属性**************************
	'Size : 每页显示
	'Page : 当前页码
	'PageCount : 结果页数
	'RecordCount : 结果记录总数
	'**********
	Public Property Get Size()
		Size = i_pSize
	End Property
	
	Public Property Get Page()
		Page = i_pNumber
	End Property
	
	Public Property Get PageCount()
		PageCount = i_pCount
	End Property
	
	Public Property Get RecordCount()
		RecordCount = i_rCount
	End Property

	'**********分页模板Top（1）******************
	'Call Header_a(Obj,SQL语句)
	'**********
	Sub Header_a(OutRs,sql)
		Set OutRs = Server.CreateObject("ADODB.Recordset")
		OutRs.open sql,i_conn,1,1
		If Not OutRs.eof Then
			i_rCount = OutRs.RecordCount
			i_pCount = Abs(Int(-(i_rCount / i_pSize)))
			If i_pNumber > i_pCount Then i_pNumber = i_pCount
			OutRs.Move i_pSize * (Abs(i_pNumber) - 1)
		End If
	End Sub
	
	'**********分页模板Top（2）******************
	'Call Header_b(Obj,数据表名,主键,查询字段,查询条件,排序)
	'SQL Server 2000
	'**********
	Sub Header_b(OutRs,sTable,sFields,sWhere,sGroup,sSort)
		Dim Cmd, sql,StrOrder, bFlag
		sSort = Trim(sSort)
		Call i_conn.Execute("select count(0) from sysobjects where id = object_id(N'["&Me.PageProcedure&"]') and OBJECTPROPERTY(id, N'IsProcedure') = 1",bFlag)
		Set Cmd = Server.CreateObject("ADODB.Command")
		With Cmd
			.ActiveConnection = i_conn
			.CommandType = 4
			If bFlag Then
				.CommandText = Me.PageProcedure
			Else
				Dim sqlParams, sqlCmd 
				sqlParams = "@TableName VARCHAR(200),@Fields VARCHAR(200),@sqlWhere VARCHAR(200) = '',@Group VARCHAR(200) = '',@OrderField VARCHAR(200) = '',@PageSize INT,@CurrentPage INT = 1"
				sqlCmd = "DECLARE @SortColumn VARCHAR(200) DECLARE @Operator CHAR(2) DECLARE @SortTable VARCHAR(200) DECLARE @SortName VARCHAR(200) Declare @totalRecord int DECLARE @totalPage INT Declare @sql nvarchar(4000) IF @Fields = '' SET @Fields = '*' IF @sqlWhere = '' SET @sqlWhere = ' WHERE 1=1' ELSE SET @sqlWhere = ' WHERE ' + @sqlWhere IF @Group <>'' SET @Group = ' GROUP BY ' + @Group IF @OrderField <> '' BEGIN DECLARE @pos1 INT, @pos2 INT SET @OrderField = REPLACE(REPLACE(@OrderField, ' asc', ' ASC'), ' desc', ' DESC') IF CHARINDEX(' DESC', @OrderField) > 0 IF CHARINDEX(' ASC', @OrderField) > 0 BEGIN IF CHARINDEX(' DESC', @OrderField) < CHARINDEX(' ASC', @OrderField) SET @Operator = '<=' ELSE SET @Operator = '>=' END ELSE SET @Operator = '<=' ELSE SET @Operator = '>=' SET @SortColumn = REPLACE(REPLACE(REPLACE(@OrderField, ' ASC', ''), ' DESC', ''), ' ', '') SET @pos1 = CHARINDEX(',', @SortColumn) IF @pos1 > 0 SET @SortColumn = SUBSTRING(@SortColumn, 1, @pos1-1) SET @pos2 = CHARINDEX('.', @SortColumn) IF @pos2 > 0 BEGIN SET @SortTable = SUBSTRING(@SortColumn, 1, @pos2-1) IF @pos1 > 0  SET @SortName = SUBSTRING(@SortColumn, @pos2+1, @pos1-@pos2-1) ELSE SET @SortName = SUBSTRING(@SortColumn, @pos2+1, LEN(@SortColumn)-@pos2) END ELSE BEGIN SET @SortTable = @TableName SET @SortName = @SortColumn END END DECLARE @type varchar(50) DECLARE @prec int SELECT @type=t.name, @prec=c.prec FROM sysobjects o  JOIN syscolumns c on o.id=c.id JOIN systypes t on c.xusertype=t.xusertype WHERE o.name = @SortTable AND c.name = @SortName IF CHARINDEX('char', @type) > 0 SET @type = @type + '(' + CAST(@prec AS varchar) + ')' DECLARE @TopRows INT SET @TopRows = @PageSize * (@CurrentPage-1) + 1 if (@SqlWhere='' or @sqlWhere=NULL) set @sql = 'select @totalRecord = count(1) from ' + @TableName + @Group else set @sql = 'select @totalRecord = count(1) from ' + @TableName + @sqlWhere + @Group EXEC sp_executesql @sql,N'@totalRecord int OUTPUT',@totalRecord OUTPUT SELECT @totalPage=CEILING((@totalRecord+0.0)/@PageSize) SELECT @totalRecord AS RecordCount,@totalPage AS PageCount EXEC(' DECLARE @SortColumnBegin ' + @type + ' SET ROWCOUNT ' + @TopRows + ' SELECT @SortColumnBegin=' + @SortColumn + ' FROM ' + @TableName + ' ' + @sqlWhere + ' ' + @Group + ' ORDER BY ' + @OrderField + ' SET ROWCOUNT ' + @PageSize + ' SELECT ' + @Fields + ' FROM ' + @TableName + ' ' + @sqlWhere + ' AND ' + @SortColumn + '' + @Operator + '@SortColumnBegin ' + @Group + ' ORDER BY ' + @OrderField)"
				.CommandText = "sp_executesql"
				.Parameters.Append .CreateParameter("@stmt",203,,8000,sqlCmd)
				.Parameters.Append .CreateParameter("@parameters",203,,8000,sqlParams)
			End If
			.Parameters.Append .CreateParameter ("@TableName",200,,1000,sTable)
			.Parameters.Append .CreateParameter ("@Fields",200,,1000,sFields)
			.Parameters.Append .CreateParameter ("@sqlWhere",200,,1000,sWhere)
			.Parameters.Append .CreateParameter ("@Group",200,,1000,sGroup)
			.Parameters.Append .CreateParameter ("@OrderField",200,,200,sSort)
			.Parameters.Append .CreateParameter ("@PageSize",5,,,i_pSize)
			.Parameters.Append .CreateParameter ("@CurrentPage",5,,,i_pNumber)
			Set OutRs = .Execute
		End With
		Set Cmd = Nothing
		i_pCount = OutRs("PageCount")
		i_rCount = OutRs("RecordCount")
		Set OutRs = OutRs.NextRecordset
		If i_pNumber > i_pCount Then i_pNumber = i_pCount
	End Sub

	'**********分页模板Top（3）******************
	'Call Header_c(Obj,数据表名,查询字段,查询条件,排序)
	'SQL Server 2005
	'**********
	Sub Header_c(OutRs,sTable,sFields,sWhere,sGroup,sSort)
		Dim Cmd, sql,StrOrder, bFlag
		sSort = Trim(sSort)
		Call i_conn.Execute("select Top 1 1 from sysobjects where id = object_id(N'["&Me.PageProcedure&"]') and OBJECTPROPERTY(id, N'IsProcedure') = 1",bFlag)
		Set Cmd = Server.CreateObject("ADODB.Command")
		With Cmd
			.ActiveConnection = i_conn
			.CommandType = 4
			If bFlag Then
				.CommandText = Me.PageProcedure
			Else
				Dim sqlParams, sqlCmd 
				sqlParams = "@TableName varchar(350),@Fields varchar(5000) = '*',@sqlWhere varchar(5000) = Null,@Group varchar(5000) = Null,@OrderField varchar(5000),@PageSize int,@CurrentPage int = 1"
				sqlCmd = "Declare @sql nvarchar(4000) Declare @totalRecord int  DECLARE @totalPage INT if (@SqlWhere='' or @sqlWhere=NULL) set @sql = 'select @totalRecord = count(1) from ' + @TableName else set @sql = 'select @totalRecord = count(1) from ' + @TableName + ' where ' + @sqlWhere EXEC sp_executesql @sql,N'@totalRecord int OUTPUT',@totalRecord OUTPUT SELECT @totalPage=CEILING((@totalRecord+0.0)/@PageSize) SELECT @totalRecord AS 'RecordCount',@totalPage AS 'PageCount' IF @Group <>'' SET @Group = ' GROUP BY ' + @Group if (@SqlWhere='' or @sqlWhere=NULL) set @sql = 'Select * FROM (select ROW_NUMBER() Over(order by ' + @OrderField + ') as rowId,' + @Fields + ' from ' + @TableName + @Group else set @sql = 'Select * FROM (select ROW_NUMBER() Over(order by ' + @OrderField + ') as rowId,' + @Fields + ' from ' + @TableName + ' where ' + @SqlWhere + @Group if @CurrentPage<=0  Set @CurrentPage = 1 if @CurrentPage>@TotalPage Set @CurrentPage = @TotalPage Declare @StartRecord int Declare @EndRecord int set @StartRecord = (@CurrentPage-1)*@PageSize + 1 set @EndRecord = @StartRecord + @PageSize - 1 set @Sql = @Sql + ') as tempTable where rowId >=' + CONVERT(VARCHAR(50),@StartRecord) + ' and rowid<= ' + CONVERT(VARCHAR(50),@EndRecord) Exec(@Sql)"
				.CommandText = "sp_executesql"
				.Parameters.Append .CreateParameter("@stmt",203,,8000,sqlCmd)
				.Parameters.Append .CreateParameter("@parameters",203,,8000,sqlParams)
			End If
			.Parameters.Append .CreateParameter ("@TableName",200,,1000,sTable)
			.Parameters.Append .CreateParameter ("@Fields",200,,1000,sFields)
			.Parameters.Append .CreateParameter ("@sqlWhere",200,,1000,sWhere)
			.Parameters.Append .CreateParameter ("@Group",200,,1000,sGroup)
			.Parameters.Append .CreateParameter ("@OrderField",200,,200,sSort)
			.Parameters.Append .CreateParameter ("@PageSize",5,,,i_pSize)
			.Parameters.Append .CreateParameter ("@CurrentPage",5,,,i_pNumber)
			Set OutRs = .Execute
		End With
		Set Cmd = Nothing
		i_pCount = OutRs("PageCount")
		i_rCount = OutRs("RecordCount")
		Set OutRs = OutRs.NextRecordset
		If i_pNumber > i_pCount Then i_pNumber = i_pCount
	End Sub

	'**********分页模板End（1）******************
	'Footer_a(后续参数,表现类型)
	'**********
	Function Footer_a(str,p_Type)
		If i_pCount = 1 Then Exit Function
		If str <> "" Then
			CurrentPath = regExpReplace(CurrentPath,"&?"&str,"",false)
			str = str & "&"
		End If
		echo "<span class=""pageIntroA"">"
		If i_rCount<>"" Then 
			echo "总数:<kbd class=""p_total"">"&i_rCount&"</kbd> "&vbCrlf
		End If
		echo "每页:<kbd class=""p_size"">"&i_pSize&"</kbd> "&vbCrlf
		echo "</span>"
		echo "<span class=""pageContorlA"">"
		If i_pNumber<=1 Then
			echo  "<a href=""#;"" class=""p_disabled"" disabled title=""已经是第一页了"">首页</a> <a href=""#;"" class=""p_disabled"" disabled>上一页</a> "
		Else
			echo "<a href="""&CurrentPath&str&"PageID=1"" class=""p_start"" title=""第一页"">首页</a> <a href="""&CurrentPath&str&"PageID="&i_pNumber-1&""" class=""p_pre"">上一页</a> "
		End If
		If i_pCount="" Then
			echo "<a href="""&CurrentPath&str&"PageID="&i_pNumber+1&""" class=""p_next"">下一页</a>  <a href="""&CurrentPath&str&"PageID=999999999"" class=""p_end"">尾页</a> "
		Else
			If i_pNumber>=i_pCount Then
				echo "<a href=""#;"" class=""p_disabled"" disabled>下一页</a> <a href=""#;"" class=""p_disabled"" disabled title=""已经是最后一页了"">尾页</a>"
			Else
				echo "<a href="""&CurrentPath&str&"PageID="&i_pNumber+1&""" class=""p_next"">下一页</a> <a href="""&CurrentPath&str&"PageID="&i_pCount&""" class=""p_end"" title=""最后一页"">尾页</a>"
			End If
		End If
		Call CommonFooterContorl(str,p_Type)
		echo "</span>"
	End Function
	
	
	'**********分页模板End（2）******************
	'Footer_b(后续参数,表现类型)
	'**********
	Function Footer_b(str,p_Type)
		Dim i, m
		m = 9
		If i_pCount = 1 Then Exit Function
		If str <> "" Then
			CurrentPath = regExpReplace(CurrentPath,"&?"&str,"",false)
			CurrentPath = Replace(CurrentPath,"?&","")
			str = str & "&"
		End If
		echo "<span class=""pageIntroB"">"
		If i_rCount <> "" Then
			echo "总数:<kbd class=""p_total"">"&i_rCount&"</kbd>"
		End If
		echo " 每页:<kbd class=""p_size"">"&i_pSize&"</kbd>"
		echo "</span>"
		echo "<span class=""pageContorlB"">"
		If i_pNumber = 1 Then 
			echo " <a href=""#;"" class=""p_disabled"" disabled title=""已经是第一页了"">上一页</a>"
		Else
			echo " <a href="""&CurrentPath&"PageID="&i_pNumber-1&""&str&""" class=""p_pre"">上一页</a> "
		End If
		If i_pNumber > m - 4 Then 
			echo " <a href="""&CurrentPath&"PageID=1"&str&""" class=""p_start"" title=""第一页"">1</a> "
			If i_pNumber > m - 3 Then echo " ... "
		End If
		For i = i_pNumber - m + 5 to i_pNumber + m - 1
			If i > 0 and i <= i_pCount Then 
				If i = i_pNumber Then
					echo " <strong class=""p_cur"">"&i&"</strong> "
				Else
					echo " <a href="""&CurrentPath&"PageID="&i&""&str&""" class=""p_page"">"&i&"</a> "
				End If
			End If
			If i_pNumber < m - 3 And i > m - 1 Then Exit For
			If i_pNumber > m - 5 And i >= i_pNumber + m - 5 Then Exit For
		Next
		If i_pNumber < i_pCount - m + 5 Then
			If i_pNumber < i_pCount - m + 4 Then echo " ... "
			echo " <a href="""&CurrentPath&"PageID="&i_pCount&""&str&""" class=""p_end"" title=""最后一页"">"&i_pCount&"</a> "
		End If
		If i_pNumber = i_pCount Then 
			echo " <a href=""#;"" class=""p_disabled"" disabled title=""已经是最后一页了"">下一页</a> "
		Else
			echo " <a href="""&CurrentPath&"PageID="&i_pNumber+1&""&str&""" class=""p_next"">下一页</a> "
		End If
		Call CommonFooterContorl(str,p_Type)
		echo "</span>"
	End Function

	'**********
	'CommonFooterContorl(后续参数,表现类型)
	'**********
	Private Sub CommonFooterContorl(str,p_Type)
		Select Case p_Type
		Case "select"
			echo "&nbsp;跳转到:<select name=""PageID"" onchange=""location.href='"&CurrentPath&"?PageID='+this.options[this.selectedIndex].value+'"&str&"'"" class=""p_select"">"&vbCrlf
			Dim i
			For i=1 to i_pCount
				echo "<option value="""&i&""""
				If i=i_pNumber Then echo " selected"
				echo ">第"&i&"页</option>"&vbCrlf
			Next
			echo "</select>"&vbCrlf
		Case Else
			Randomize
			Dim PageID : PageID = "PageID" & Int(Rnd() * 10000000)
			echo " 跳转到:<input type=""text"" id="""&PageID&""" name=""PageID"" onkeydown=""if(event.keyCode==13) document.getElementById('btn_"&PageID&"').click();"" size=""3"" value="""&i_pNumber&""" onclick=""this.select()"" maxlength=8 class=""p_text""> "&vbCrlf & _
							 "<input type=""button"" value=""GO"" onclick=""location.href='"&CurrentPath&"?PageID='+document.getElementById('"&PageID&"').value+'"&str&"'"" id=""btn_"&PageID&""" class=""p_btn"">"
		End Select
	End Sub

	Private Function regExpReplace(ByVal str,re,restr,isCase)	'内容,正则
		If Len(str) > 0 Then
			Dim Obj
			Set Obj = New Regexp
			With Obj
				If isCase Then .IgnoreCase = False Else .IgnoreCase = True 
				.Global = True
				.Pattern = re
				regExpReplace = .Replace(str,restr)
			End With
			Set Obj = Nothing
		End If
	End Function

	Private Sub echo(str)
		Response.Write str
	End Sub
End Class
%>
