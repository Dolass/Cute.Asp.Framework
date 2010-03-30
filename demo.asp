<!--#include file="config.inc.asp"-->
<!--#include file="include/helper/json.asp"-->
<%
Casp.params a
Casp.params b
b.Add "b1",2
a.add "a",1
a.add "b",b
die  Casp.json.toJSON("aa", a, false)
%>