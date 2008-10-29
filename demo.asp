<!--#include file="config.inc.asp"-->
<!--#include file="include/library/file.asp"-->
<!--#include file="include/helper/export.asp"-->
<!--#include file="include/helper/des.asp"-->
<%
Tpub.db.Open
'Tpub.Export.Conn = Tpub.db.Conn
'Call Tpub.export.txt("select top 10 book_name from list_book ","")
'Tpub.Des.Key = "ss"
'echo Tpub.Des.Encode("admin")&vbCrlf
'echo Tpub.Des.Decode(Tpub.Des.Encode("admin"))
Function getBookObj(id)
	Dim sql
	sql = "select Top 10 * from List_book where id="&id
	Set getBookObj = Tpub.db.GetRowObject(sql)
End Function

Dim Book5
Set Book5 = getBookObj(5)
echo Book5("id")
%>