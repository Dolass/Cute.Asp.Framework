<%
'**********
'	class		: Wrapper Class
'	File Name	: class.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-6-28
'**********


'**********
'	示例
'**********
'**********
'	构建类
'**********
Dim Tpub
Set Tpub = New Class_Wrap


Class Class_Wrap

	Public Db,[String],Params,[Array],Upload,Page,File,Debug,Cookie,Session,[Date]

	Public SHA1,Md5,Json,ValidCode,Export,Email,InterFace,BEDcode,Xml

	Private Sub Class_Initialize()
		On Error Resume Next
		Set Db				= New Class_Db				'数据库操作类
		Set [String]		= New Class_String			'String操作类
		Set Params			= New Class_Params			'Dictionary简化操作类
		Set [Array]			= New Class_Array			'数组操作类
		Set Upload			= New Class_Upload			'上传类
		Set Page			= New Class_Page			'分页类
		Set File			= New Class_File			'文件操作类
		Set Debug			= New Class_Debug			'Debug工具类
		Set Cookie			= New Class_Cookie			'Cookies操作类
		Set Session			= New Class_Session			'Session操作类
		Set [Date]			= New Class_Date			'Date操作类

		Set SHA1			= Class_SHA1()				'SHA1编码
		Set Md5				= New Class_Md5				'Md5加密
		Set Json			= Class_Json()				'Json操作类
		Set ValidCode		= New Class_ValidCode		'验证码
		Set Export			= New Class_Export			'Export Data
		Set Email			= New Class_Email			'Email发送类
		Set InterFace		= New Class_Interface		'远程获取类
		Set BEDcode			= New Class_BasicEncode		'简单加密解密
		Set Xml				= New Class_XML				'XML操作类
		Err.Clear
		On Error Goto 0
	End Sub

	Private Sub Class_Terminate()
		If IsObject(Db) Then			Set Db				= Nothing
		If IsObject([String]) Then		Set [String]		= Nothing
		If IsObject(Params) Then		Set Params			= Nothing
		If IsObject([Array]) Then		Set [Array]			= Nothing
		If IsObject(Upload) Then		Set Upload			= Nothing
		If IsObject(Page) Then			Set Page			= Nothing
		If IsObject(File) Then			Set File			= Nothing
		If IsObject(Debug) Then			Set Debug			= Nothing
		If IsObject(Cookie) Then		Set Cookie			= Nothing
		If IsObject(Session) Then		Set Session			= Nothing
		If IsObject([Date]) Then		Set [Date]			= Nothing

		If IsObject(SHA1) Then			Set SHA1			= Nothing
		If IsObject(Md5) Then			Set Md5				= Nothing
		If IsObject(Json) Then			Set Json			= Nothing
		If IsObject(ValidCode) Then		Set ValidCode		= Nothing
		If IsObject(Export) Then		Set Export			= Nothing
		If IsObject(Email) Then			Set Email			= Nothing
		If IsObject(InterFace) Then		Set InterFace		= Nothing
		If IsObject(BEDcode) Then		Set BEDcode			= Nothing
		If IsObject(Xml) Then			Set Xml				= Nothing
	End Sub

	%>
	<!--#include file="library/common.asp"-->
	<%
End Class
%>
<script language="vbscript" runat="server">
	Set Tpub = Nothing	'页面执行完毕后自动释放对象
</script>