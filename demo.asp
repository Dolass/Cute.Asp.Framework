<!--#include file="config.inc.asp"-->
<!--#include file="include/helper/export.asp"-->
<!--#include file="include/helper/des.asp"-->
<%
Tpub.db.Open
Tpub.Export.Conn = Tpub.db.Conn
Call Tpub.export.txt("select top 10 book_name from list_book ","")
echo Tpub.Des.Encode("ss","admin")&vbCrlf
echo Tpub.Des.Decode("ss",Tpub.Des.Encode("ss","admin"))
%>