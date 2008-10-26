<%
'**********
'	class		: provide a simple way to parse JSON
'	File Name	: json.asp
'	Version		: 0.1.0
'	Updater		: TerranC
'	Date		: 2008-5-21
'**********

Class JSONClass
	' 定义类属性，默认为Private
	Private SqlString ' 用于设置Select
	Private JSON ' 返回的JSON对象的名称
	Private DBConnection ' 连接到数据库的Connection对象

	' 可以外部调用的公共方法

	Public Function GetJSON()
		Dim Rs
		Dim returnStr
		Dim i
		Dim oneRecord

		' 获取数据
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.Open SqlString, DBConnection, 1, 1
		' 生成JSON字符串
		If Rs.EOF = false And Rs.Bof = false Then
			returnStr = "{ "& JSON & ":{ records:["
			While Rs.EOF = false
				' -------
				oneRecord = "{"
				For i = 0 To Rs.Fields.Count -1
					oneRecord = oneRecord & Chr(34) &Rs.Fields(i).Name&Chr(34)&":"
					oneRecord = oneRecord & Chr(34) &Rs.Fields(i).Value&Chr(34) &","
				Next
				'去除记录最后一个字段后的","
				oneRecord = Left(oneRecord, InStrRev(oneRecord, ",") -1)
				oneRecord = oneRecord & "},"
				'------------
				returnStr = returnStr & oneRecord
				Rs.MoveNext
			Wend
			' 去除所有记录数组后的","
			returnStr = Left(returnStr, InStrRev(returnStr, ",") -1)
			returnStr = returnStr & "]}}"
		End If
		Rs.Close
		Set Rs = Nothing
		GetJSON = returnStr
	End Function
End Class
%>