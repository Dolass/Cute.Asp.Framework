<!--#include file="config.inc.asp"-->
<!--#include file="include/library/file.asp"-->
<!--#include file="include/helper/export.asp"-->
<!--#include file="include/helper/des.asp"-->
<%
'Tpub.db.Open
'Tpub.Export.Conn = Tpub.db.Conn
'Call Tpub.export.txt("select top 10 book_name from list_book ","")
'Tpub.Des.Key = "ss"
'echo Tpub.Des.Encode("admin")&vbCrlf
'echo Tpub.Des.Decode(Tpub.Des.Encode("admin"))

echo Tpub.file.GetFolderSize("js")
%>