<!--#include file="config.inc.asp"-->
<!--#include file="include/helper/json.asp"-->
<%
Tpub.db.Open
Tpub.db.Exec rs,"select Top 10 * from CommendSetup_User"
echo Tpub.Json.GetJson("data",rs)
%>