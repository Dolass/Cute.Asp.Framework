<!--#include file="config.inc.asp"-->
<!--#include file="include/helper/des.asp"-->
<!--#include file="include/helper/md5.asp"-->
<%
Tpub.Des.Key = "12345678"
echo Tpub.Des.base64encode(Tpub.Des.encode("ss"))
echo "<br>"
echo Tpub.Des.decode(Tpub.Des.base64decode("NTZiNTg3OTA5MmU1MTJiNA=="))
%>