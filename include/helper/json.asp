<%
'**********
'	class		: provide a simple way to parse JSON
'	File Name	: json.asp
'	Version		: 0.1.0
'	Updater		: TerranC
'	Date		: 2008-5-21
'**********

Class Class_Json

	'**********
    ' 函数名: GetJSON
    ' 参　数: name	-- 对象名
    ' 参　数: obj	-- Recordset对象
    ' 作  用: 把Recordset对象转换成JSON字符串
    '**********
	Public Function GetJSON(name,obj)
		Dim Rs
		Dim returnStr
		Dim i
		Dim oneRecord

		' 获取数据
		Set Rs = obj
		' 生成JSON字符串
		If obj.EOF = false And Rs.Bof = false Then
			returnStr = "{ """& name & """:{ ""records"":["
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