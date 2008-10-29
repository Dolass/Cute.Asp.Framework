<%
Option Explicit
Response.Buffer = True
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
Tpub.db.ServerIp = "(local)"
Tpub.db.Database = "qds8data"
Tpub.db.UserName = "sa"
Tpub.db.Password = "sqladmin"

Sub Finish()
    Dim RunTime : RunTime = round((Timer() - StartTime), 3)
    echo "<p style=""text-align:center;font-size:11px;"">Page rendered in " & RunTime & " seconds.</p>"
End Sub
%>
