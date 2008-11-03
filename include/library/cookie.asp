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
Class Class_Cookie
	Public	Mark	'前缀

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
    ' 作  用: Add a cookie
    '**********
	Sub [set](Key, Value, Options)
        Response.Cookies(Me.Mark & Key) = Value
        If Not (IsNull(Options) Or IsEmpty(Options) Or Options = "") Then
            If IsArray(Options) Then
                Dim l : l = UBound(Options)
                Response.Cookies(Me.Mark & Key).Expire = Options(0)
                If l = 1 Then Response.Cookies(Me.Mark & Key).Path = Options(1)
                If l = 2 Then Response.Cookies(Me.Mark & Key).Domain = Options(2)
            Else
                Response.Cookies(Me.Mark & Key).Expire = Options
            End If
        End If
    End Sub

    '**********
    ' 函数名: get
    ' 作  用: get a cookies
    '**********
	Function [get](Key)
        [get] = Request.Cookies(Me.Mark & Key)
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: Remove a cookie
    '**********
	Sub Remove(Key)
         Response.Cookies(Me.Mark & Key) = Empty
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: Remove all cookies
    '**********
	Sub RemoveAll()
        Clear()
    End Sub

    '**********
    ' 函数名: Clear
    ' 作  用: Remove all cookies
    '**********
	Private Sub Clear()
        Dim iCookie
        For Each iCookie In Request.Cookies
            Response.Cookies(iCookie).Expires = formatDateTime(Now)
        Next
    End Sub

End Class
%>
