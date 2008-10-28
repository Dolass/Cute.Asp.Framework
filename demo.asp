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
	Set getBookObj = GetEntryObject(sql)
End Function

Function GetEntryObject(sql)
	Tpub.db.Exec rs,sql
	If rs.eof Then 
		Set GetEntryObject = Nothing
		Exit Function
	End If
	Tpub.Params.Open
	For i=0 To rs.fields.count-1
		If IsNumeric(rs(i)) Then
			Tpub.Params.Add rs.Fields(i).Name,rs(i)+0
		Else
			Tpub.Params.Add rs.Fields(i).Name,rs(i)&""
		End if
	Next
	Tpub.db.closeRs(rs)
	Set GetEntryObject = Tpub.Params.Item
	Tpub.Params.Close
End Function


Set Book5 = getBookObj(5)
Set Book6 = getBookObj(6)
echo Book5("id")
%>