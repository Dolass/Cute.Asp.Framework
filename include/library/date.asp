<%
'**********
'	class		: A Date class
'	File Name	: Date.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-4-2
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********
Class Class_Date
	Public TimeZone
	
    '**********
    ' 函数名: class_Initialize
    ' 作  用: Save the session
    '**********
	Private Sub class_initialize()
		TimeZone = 8
    End Sub

    '**********
    ' 函数名: class_Terminate
    ' 作  用: Deconstrutor
    '**********
	Private Sub class_Terminate()
    End Sub

	Private Function getMistiming(sDate)
		getMistiming = DateDiff("s","1970-1-1 00:00:00",DateAdd("h",Me.TimeZone,sDate))
	End Function

    '**********
    ' 函数名: toGMTdate
    ' 参  数: sDate
    ' 作  用: 获取GMT时间
    '**********
	Function toGMTdate(sDate)
	  Dim dWeek,dMonth
	  Dim strZero,strZone
	  strZero="00"
	  If Me.TimeZone > 0 Then
	  	strZone = "+"&Right("0"&Me.TimeZone,2)&"00"
	  Else
	  	strZone = "-"&Right("0"&Me.TimeZone,2)&"00"
	  End If
	  dWeek=Array("Sun","Mon","Tue","Wes","Thu","Fri","Sat")
	  dMonth=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
	  toGMTdate = dWeek(WeekDay(sDate)-1)&", "&Right(strZero&Day(sDate),2)&" "&dMonth(Month(sDate)-1)&" "&Year(sDate)&" "&Right(strZero&Hour(sDate),2)&":"&Right(strZero&Minute(sDate),2)&":"&Right(strZero&Second(sDate),2)&" "&strZone
	End Function

    '**********
    ' 函数名: toUnixEpoch
    ' 参  数: sDate
    ' 作  用: 获取时间戳
    '**********
	Function toUnixEpoch(sDate)
		toUnixEpoch = DateDiff("s", "1970-1-1 00:00:00", sDate) -  getMistiming("1970-1-1 00:00:00")
	End Function

    '**********
    ' 函数名: fromUnixEpoch
    ' 参  数: iNumber	--  时间戳
    ' 作  用: 获取当地时间
    '**********
	Function fromUnixEpoch(iNumber)
		fromUnixEpoch = DateAdd("s",iNumber +  getMistiming("1970-1-1 00:00:00"),"1970-1-1 00:00:00")
	End Function

	'**********
	' 函数名: zodiac
	' 参  数: bYear as birthday year
	' 作  用: 计算所属生肖
	'**********
	Function zodiac(bYear)
		If bYear > 0 Then
			Dim zodiacList : zodiacList = Array("猴", "鸡", "狗", "猪", "鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊")
			zodiac = zodiacList(bYear mod 12)
		End If
	End Function

	'**********
	' 函数名: Constellation
	' 参  数: Birth as birthday
	' 作  用: 计算所属生肖
	'**********    
	Function Constellation(Birth)
        If Year(Birth) <1951 Or Year(Birth) > 2049 Then Exit Function
        Dim BirthDay, BirthMonth
        BirthDay = Day(Birth)
        BirthMonth = Month(Birth)
        Select Case BirthMonth
            Case 1
                If BirthDay>= 21 Then
                    Constellation = Constellation & "水瓶"
                Else
                    Constellation = Constellation & "魔羯"
                End If
            Case 2
                If BirthDay>= 20 Then
                    Constellation = Constellation & "双鱼"
                Else
                    Constellation = Constellation & "水瓶"
                End If
            Case 3
                If BirthDay>= 21 Then
                    Constellation = Constellation & "白羊"
                Else
                    Constellation = Constellation & "双鱼"
                End If
            Case 4
                If BirthDay>= 21 Then
                    Constellation = Constellation & "金牛"
                Else
                    Constellation = Constellation & "白羊"
                End If
            Case 5
                If BirthDay>= 22 Then
                    Constellation = Constellation & "双子"
                Else
                    Constellation = Constellation & "金牛"
                End If
            Case 6
                If BirthDay>= 22 Then
                    Constellation = Constellation & "巨蟹"
                Else
                    Constellation = Constellation & "双子"
                End If
            Case 7
                If BirthDay>= 23 Then
                    Constellation = Constellation & "狮子"
                Else
                    Constellation = Constellation & "巨蟹"
                End If
            Case 8
                If BirthDay>= 24 Then
                    Constellation = Constellation & "处女"
                Else
                    Constellation = Constellation & "狮子"
                End If
            Case 9
                If BirthDay>= 24 Then
                    Constellation = Constellation & "天秤"
                Else
                    Constellation = Constellation & "处女"
                End If
            Case 10
                If BirthDay>= 24 Then
                    Constellation = Constellation & "天蝎"
                Else
                    Constellation = Constellation & "天秤"
                End If
            Case 11
                If BirthDay>= 23 Then
                    Constellation = Constellation & "射手"
                Else
                    Constellation = Constellation & "天蝎"
                End If
            Case 12
                If BirthDay>= 22 Then
                    Constellation = Constellation & "魔羯"
                Else
                    Constellation = Constellation & "射手"
                End If
            Case Else
                Constellation = ""
        End Select
    End Function

End Class
%>
