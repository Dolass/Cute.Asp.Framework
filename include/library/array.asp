<%
'**********
'	class		: Array extensive class
'	File Name	: Array.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-4-20
'**********


'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Array

	'**********
	' 函数名: max
	' 参  数: arr as a Array
	' 作  用: max — 取出最大值
	'**********
	Function max(arr)
		Dim tmp
		tmp = rsort(arr)
		max = tmp(0)
	End Function
	
	'**********
	' 函数名: min
	' 参  数: arr as a Array
	' 作  用: min — 取出最小值
	'**********
	Function min(arr)
		Dim tmp
		tmp = sort(arr)
		min = tmp(0)
	End Function
	
	'**********
	' 函数名: unshift
	' 参  数: arr as an Array
	' 作  用: unshift — 从前压入元素
	'**********
	Function unshift(arr, var)
		Dim i, tmp
		tmp = toString(arr)
		tmp = var & "," & tmp
		tmp = toArray(tmp)
		unshift = tmp
	End Function
	
	'**********
	' 函数名: shift
	' 参  数: arr as an Array
	' 作  用: shift — 从前删除元素
	'**********
	Function shift(arr)
		Dim i, tmp
		tmp = ""
		For i = 0 To UBound(arr)
			If i<>0 Then tmp = tmp & arr(i) & ","
		Next
		tmp = strip(tmp)
		shift = tmp
	End Function
	
	'**********
	' 函数名: push
	' 参  数: arr as an Array
	' 参  数: var as a variable added to an array
	' 作  用: push — 从后压入元素
	'**********
	Function push(arr, var)
		Dim tmp : tmp = toString(arr)
		tmp = tmp & "," & convComma(var)
		tmp = toArray(tmp)
		push = tmp
	End Function
	
	'**********
	' 函数名: pop
	' 参  数: arr as an array
	' 作  用: pop — 从后删除元素
	'**********
	Function pop(arr)
		Dim i, tmp
		For i = 0 To UBound(arr)
			If i<>UBound(arr) Then tmp = tmp & arr(i) & ","
		Next
		tmp = strip(tmp)
		pop = tmp
	End Function
	
	'**********
	' 函数名: strip
	' 参  数: str as a string such as "1,2,3,"
	' 作  用: Strip "," of string
	'**********
	Function strip(Str)
		If IsArray(Str) Then Str = toString(Str)
		If Left(Str, 1) = "," Then Str = Right(Str, Len(Str) -1)
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str) -1)
		Str = toArray(Str)
		strip = Str
	End Function
	
	'**********
	' 函数名: walk
	' 参  数: arr as an Array
	' 参  数: callback as callback function
	' 作  用: walk — 对数组内元素执行函数后返回新数组
	'**********
	Function walk(arr, callback)
		Dim e : e = ""
		Dim tmp : tmp = ""
		For Each e in arr
			If IsArray(e) Then
				Execute("tmp=tmp&" & callback & "(""" & toString(e) & """)" & "&"",""")
			Else
				Execute("tmp=tmp&" & callback & "(""" & e & """)" & "&"",""")
			End If
		Next
		tmp = strip(tmp)
		walk = tmp
	End Function
	
	'**********
	' 函数名: splice
	' 参  数: arr as an array
	' 参  数: start as start index
	' 参  数: final as end index
	' 作  用: splice — 从一个数组中移除一个或多个元素
	'**********
	Function splice(arr, start, final)
		Dim i, temp, tmp
		If start > final Then
			temp = start
			start = final
			final = temp
		End If
		For i = 0 To UBound(arr)
			If i < start Or i > final Then tmp = tmp & arr(i) & ","
		Next
		tmp = strip(tmp)
		splice = tmp
	End Function
	
	'**********
	' 函数名: fill
	' 参  数: arr as a Array
	' 参  数: index as index to insert into an array
	' 参  数: value as element to insert into an array
	' 作  用: fill — 插入元素
	'**********
	Function fill(arr, index, Value)
		Dim i, tmp
		For i = 0 To UBound(arr)
			If i <> index Then
				tmp = tmp & arr(i) & ","
			Else
				tmp = tmp & Value & "," & arr(i) & ","
			End If
		Next
		tmp = strip(tmp)
		fill = tmp
	End Function
	
	'**********
	' 函数名: unique
	' 参  数: arr as a Array
	' 作  用: unique — 移除重复的元素
	'**********
	Function unique(arr)
		Dim tmp, e
		For Each e in arr
			If InStr(1, tmp, e) = 0 Then
				tmp = tmp & e & ","
			End If
		Next
		tmp = strip(tmp)
		unique = tmp
	End Function

	'**********
	' 函数名: reverse
	' 参  数: arr as a Array
	' 作  用: reverse — 反向
	'**********
	Function reverse(arr)
		Dim tmp, e
		For Each e in arr
			tmp = tmp & e & ","
		Next
		tmp = StrReverse(tmp)
		tmp = strip(tmp)
		reverse = tmp
	End Function
	
	'**********
	' 函数名: search
	' 参  数: arr as a Array
	' 参  数: value as searching value
	' 作  用: search — 查询元素，不存在则返回False
	'**********
	Function search(arr, Value)
		Dim a_tmp
		a_tmp = Filter(arr,Value)
		If UBound(a_tmp) = -1 Then
			search = False
			Exit Function
		End If
		search = a_tmp(0)
	End Function
	
	'**********
	' 函数名: rand
	' 参  数: arr as a Array
	' 参  数: num as specifies how many entries you want to pick
	' 作  用: rand — 乱序
	'**********
	Function rand(arr, num)
		Dim tmpi, tmp, i
		For i = 0 To num -1
			Randomize
			tmpi = Int((UBound(arr) + 1) * Rnd)
			tmp = tmp & arr(tmpi) & ","
		Next
		tmp = strip(tmp)
		rand = tmp
	End Function
	
	'**********
	' 函数名: sort
	' 参  数: arr as a Array
	' 作  用: sort — 顺序
	'**********
	Function sort(arr)
		Dim tmp, i, j
		ReDim tmpA(UBound(arr))
		For i = 0 To UBound(tmpA)
			tmpA(i) = CDbl(arr(i))
		Next
		For i = 0 To UBound(tmpA)
			For j = i + 1 To UBound(tmpA)
				If tmpA(i) > tmpA(j) Then
					tmp = tmpA(i)
					tmpA(i) = tmpA(j)
					tmpA(j) = tmp
				End If
			Next
		Next
		sort = tmpA
	End Function
	
	'**********
	' 函数名: rsort
	' 参  数: arr as a Array
	' 作  用: rsort — 倒序
	'**********
	Function rsort(arr)
		Dim tmp, i, j
		ReDim tmpA(UBound(arr))
		For i = 0 To UBound(tmpA)
			tmpA(i) = CDbl(arr(i))
		Next
		For i = 0 To UBound(tmpA)
			For j = i + 1 To UBound(tmpA)
				If tmpA(i) < tmpA(j) Then
					tmp = tmpA(i)
					tmpA(i) = tmpA(j)
					tmpA(j) = tmp
				End If
			Next
		Next
		rsort = tmpA
	End Function

	'**********
	' 函数名: shuffle
	' 参  数: arr as a Array
	' 作  用: shuffle — 随机排序
	'**********
	Function shuffle(arr)
		Dim m, n, i
		'i = search(arr,rand(arr,1))
		'arr = splice(arr,i,i+1)
		Randomize   
		For i = 0 to UBound(arr)
			m = Int(Rnd()*i)
			n = arr(m)
			arr(m) = arr(i) 
			arr(i) = n
		Next
		shuffle = arr
	End Function

	'**********
	' 函数名: convComma
	' 参  数: star as a string
	'**********
	Function convComma(str)
		convComma = Replace(str,",","&#44;")
	End Function

	'**********
	' 函数名: implode
	' 参  数: glue as a split character
	' 参  数: arr as a output array
	' 作  用: Join array elements with a string
	'**********
	Function toString(arr)
		If IsArray(arr) Then
			Dim tmp
			tmp = Join(arr, ",")
			toString = tmp
		Else
			toString = arr
		End If
	End Function

	'**********
	' 函数名: toArray
	' 参  数: str as a string converted to an array
	' 作  用: Convert to an array
	' Remarks: dim a : a = "a, b, c"
	'		   prinr(toArray(a))
	'**********
	Function toArray(str)
		Dim tmp
		tmp = Split(str, ",")
		toArray = tmp
	End Function

End Class

%>