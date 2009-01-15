<!--#include file="config.inc.asp"-->
<!--#include file="include/library/cache.asp"-->
<!--#include file="include/helper/export.asp"-->
<!--#include file="include/helper/des.asp"-->
<!--#include file="include/helper/md5.asp"-->
<%
echo Left(Tpub.md5.encode("0123456789ABCDEFFEDCBA98",32),24)
%>