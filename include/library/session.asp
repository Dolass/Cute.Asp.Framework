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
        Session(Me.Mark & Key) = Value
    End Sub

    '**********
    ' 函数名: get
    ' 作  用: get a Session
    '**********
	Function [get](Key)
        [get] = Session(Me.Mark & Key)
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: Remove a Session
    '**********
	Sub Remove(Key)
        Session.Contents.Remove(Me.Mark & Key)
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: Remove all Session
    '**********
	Sub RemoveAll()
		Session.abandon()
	End Sub

    '**********
    ' 函数名: Clear
    ' 作  用: Remove all cookies
    '**********
	Private Sub Clear()
        Dim iSession
        For Each iSession In Session.Contents
            Session.Contents.Remove(iSession)
        Next
    End Sub
End Class
%>
