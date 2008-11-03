<%
'**********
'	class		: A Caching class
'	File Name	: Cache.asp
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
Class Class_Session
	Public	Mark	'前缀

    Public Property Let Timeout(Value)
		If IsNumeric(Value) Then Session.Timeout = Value
    End Property

    Public Default Property Get Contents(Value)
        contents = [get](Value)
    End Property

    '**********
    ' 函数名: class_Initialize
    ' 作  用: Save the session
    '**********
	Private Sub class_initialize()
		Mark = "cute_"
    End Sub

    '**********
    ' 函数名: class_Terminate
    ' 作  用: Deconstrutor
    '**********
	Private Sub class_Terminate()
    End Sub

    '**********
    ' 函数名: set
    ' 作  用: Add a Session
    '**********
	Sub [set](Key, Value)
		If IsObject(Value) Then
			Set Session(Me.Mark & Key) = Value
		Else
			Session(Me.Mark & Key) = Value
		End If
    End Sub

    '**********
    ' 函数名: get
    ' 作  用: get a Session
    '**********
	Function [get](Key)
		If IsObject(Session(Me.Mark & Key)) Then
			Set [get] = Session(Me.Mark & Key)
		Else
			[get] = Session(Me.Mark & Key)
		End If
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: Remove a Session
    '**********
	Sub Remove(Key)
		If IsObject(Session(Me.Mark & Key)) Then
			Set Session(Me.Mark & Key) = Nothing
		End If
        Session.Contents.Remove(Me.Mark & Key)
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: Remove all Session
    '**********
	Sub RemoveAll()
        Dim iSession
        For Each iSession In Session.Contents
			Me.Remove(iSession)
        Next
	End Sub
End Class
%>
