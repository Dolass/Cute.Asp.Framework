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
Class Class_Cache
	Public	Mark	'前缀

	Private IExpires

    Public Default Property Get Contents(Value)
        Contents = [get](Value)
    End Property

    Private Property Let Expires(Value)
        IExpires = DateAdd("n", Value, Now)	'分钟
    End Property

    Private Property Get Expires()
        Expires = IExpires
    End Property

    '**********
    ' 函数名: class_Initialize
    ' 作  用: Constructor
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
    ' 函数名: lock
    ' 作  用: lock the applaction
    '**********
	Sub Lock()
        Application.Lock()
    End Sub

    '**********
    ' 函数名: unlock
    ' 作  用: unLock the applaction
    '**********
	Sub unlock()
        Application.unLock()
    End Sub

    '**********
    ' 函数名: SetCache
    ' 作  用: Set a cache
    '**********
	Sub [set](Key, Value, Expire)
        Expires = Expire
        Lock
		Application(Mark & "_" & Key) = Value
        Application(Mark & "_" & Key & "_Expires") = Expires
        unLock
    End Sub

    '**********
    ' 函数名: getCache
    ' 作  用: Get a cache
    '**********
	Function [get](Key)
        Dim Expire
        Expire = Application(Mark & "_" & Key & "_Expires")
        If IsNull(Expire) Or IsEmpty(Expire) Then
            [get] = ""
        Else
            If IsDate(Expire) And CDate(Expire) > Now Then
                [get] = Application(Mark & "_" & Key)
            Else
                Call Remove(Mark & "_" & Key)
                Value = ""
            End If
        End If
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: remove a cache
    '**********
	Sub Remove(Key)
        Lock
        Application.Contents.Remove(Mark & "_" & Key)
        Application.Contents.Remove(Mark & "_" & Key & "_Expires")
        unLock
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: remove all cache
    '**********
	Sub RemoveAll()
        Lock
        Application.Contents.RemoveAll()
        unLock
    End Sub

    '**********
    ' 函数名: compare
    ' 作  用: Compare two caches
    '**********
	Function Compare(Key1, Key2)
        Dim Cache1
        Cache1 = getCache(Key1)
        Dim Cache2
        Cache2 = getCache(Key2)
        If TypeName(Cache1) <> TypeName(Cache2) Then
            Compare = True
        Else
            If TypeName(Cache1) = "Object" Then
                Compare = (Cache1 Is Cache2)
            Else
                If TypeName(Cache1) = "Variant()" Then
                    Compare = (Join(Cache1, "^") = Join(Cache2, "^"))
                Else
                    Compare = (Cache1 = Cache2)
                End If
            End If
        End If
    End Function
End Class
%>