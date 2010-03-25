<%
'-----------------------------------------------------
' 描述: Asp打包类
' 作者: 小灰(quxiaohui_0@163.com)
' 链接: http://asp2004.net http://blog.csdn.net/iuhxq http://bbs.asp2004.net
' 版本: 1.0 Beta
' 版权: 本作品可免费使用，但是请勿移除版权信息
'-----------------------------------------------------

Class Rar
    Dim Files, packname, s, s1, s2, rootpath, fso, f, buf

    Private Sub Class_Initialize
        Randomize
        Dim ranNum
        ranNum = Int(90000 * Rnd) + 10000
        packname = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum&".asp2004"
        rootpath = Server.MapPath("./")
        Set Files = CreateObject("Scripting.Dictionary")
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set s = CreateObject("ADODB.Stream")
		s.Open
        s.Type = 1
        Set s1 = CreateObject("ADODB.Stream")
        s1.Open
        s1.Type = 1
        Set s2 = CreateObject("ADODB.Stream")
        s2.Open
        s2.Type = 2
    End Sub


    Private Sub Class_Terminate
        s.Close
        Set s = Nothing
        s1.Close
        Set s1 = Nothing
        s2.Close
        Set s2 = Nothing
        Set fso = Nothing
    End Sub


    Public Sub Add(obj)
        If fso.FileExists(obj) Then
            Set f = fso.GetFile(obj)
            Files.Add obj, f.Size
        ElseIf fso.FolderExists(obj) Then
            Files.Add obj, -1
            Set f = fso.GetFolder(obj)
            Set fc = f.Files
            For Each f1 in fc
                Add(LCase(f1.Path))
            Next
        End If
    End Sub


    Public Sub Pack
        Dim Str
        a = Files.Keys
        b = Files.Items
        For i = 0 To Files.Count -1
            If b(i)>= 0 Then
                s.LoadFromFile(a(i))
                buf = s.Read
                If Not IsNull(buf) Then s1.Write(buf)
            End If
            Str = Str & b(i)&">"&Replace(a(i), rootpath, "")&vbCrLf
        Next
        Str = CStr(Right("000000000"&Len(Str), 10)) & Str
        buf = TextToStream(Str)
        s.Position = 0
        s.Write buf
        s1.Position = 0
        s.Write s1.Read
        s.SetEOS
        s.SaveToFile(packname)
    End Sub


    Public Sub UnPack
        If Not fso.FolderExists(rootpath) Then
            fso.CreateFolder(rootpath)
        End If
        Dim Size
        '转换文件大小
        s.LoadFromFile(packname)
        Size = CInt(StreamToText(s.Read(10)))
        Str = StreamToText(s.Read(Size))
        arr = Split(Str, vbCrLf)
        For i = 0 To UBound(arr) -1
            arrFile = Split(arr(i), ">")
            If arrFile(0) < 0 Then
                If Not fso.FolderExists(rootpath&arrFile(1)) Then
                    fso.CreateFolder(rootpath&arrFile(1))
                End If
            ElseIf arrFile(0) >= 0 Then
                If fso.FileExists(rootpath&arrFile(1)) Then
                    fso.DeleteFile(rootpath&arrFile(1))
                End If
                s1.Position = 0
                buf = s.Read(arrFile(0))
                If Not IsNull(buf) Then s1.Write(buf)
                s1.SetEOS
                s1.SaveToFile(rootpath&arrFile(1))
            End If
        Next
    End Sub


    Public Function StreamToText(stream)
        If IsNull(stream) Then
            StreamToText = ""
        Else
            Set sm = CreateObject("ADODB.Stream")
            sm.Open
            sm.Type = 1
            sm.Write(stream)
            sm.Position = 0
            sm.Type = 2
            sm.charset = "gb2312"
            sm.Position = 0
            StreamToText = sm.ReadText()
            sm.Close
            Set sm = Nothing
        End If
    End Function


    Public Function TextToStream(text)
        If text = "" Then
            TextToStream = Null '这里该如何写？空流？
        Else
            Set sm = CreateObject("ADODB.Stream")
            sm.Open
            sm.Type = 2
            sm.charset = "gb2312"
            sm.WriteText(text)
            sm.Position = 0
            sm.Type = 1
            sm.Position = 0
            TextToStream = sm.Read
            sm.Close
            Set sm = Nothing
        End If
    End Function

End Class
%>
