<%
'Option Explicit
Response.Buffer = True
Session.CodePage = 936
Response.Charset = "gb2312"
Dim StartTime : StartTime = timer()
%>
<!--#include file="include/ext.asp"-->
<!--#include file="include/class.asp"-->
<!--#include file="include/library/db.asp"-->
<!--#include file="include/library/date.asp"-->
<!--#include file="include/library/array.asp"-->
<!--#include file="include/library/string.asp"-->
<!--#include file="include/library/params.asp"-->
<!--#include file="include/library/session.asp"-->
<%
Tpub.db.ConnectionType = "MSSQL"
Tpub.db.ServerIp = "localhost"
Tpub.db.Database = "dataname"
Tpub.db.UserName = "sa"
Tpub.db.Password = "sa"

On Error Resume Next
Tpub.Cookie.Mark = "cute_"		'设置Cookie名称前缀
Tpub.Cache.Mark = "cute_"		'设置缓存名称前缀
Tpub.Ubb.Mode = 0				'使用基本UBB
Tpub.Date.TimeZone = 8			'设置所在时区
On Error Goto 0

Sub Finish()
    Dim RunTime : RunTime = round((Timer() - StartTime), 3)
    echo "<p style=""text-align:center;font-size:11px;"">Page rendered in " & RunTime & " seconds.</p>"
End Sub
%>
