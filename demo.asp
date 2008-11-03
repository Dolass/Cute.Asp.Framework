<!--#include file="config.inc.asp"-->
<!--#include file="include/library/cache.asp"-->
<!--#include file="include/helper/export.asp"-->
<!--#include file="include/helper/des.asp"-->
<%
Call Tpub.Session.Set("a",222)
Call Tpub.Session.Set("b",11)

Tpub.Session.RemoveAll
die Tpub.Session.Get("a")
%>