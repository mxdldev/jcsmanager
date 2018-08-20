<%
'**************************************************************
'-- File Name: Cls_Public.Asp 主函数类
'-- Version: 1.0 Beta build 20090310
'-- Date: 2009-3-10
'-- Script Written By Jiangtao
'**************************************************************

Class Cls_Public
Public UserInfo,UserID,UserName,UserPass,SQLQueryNum,UserClassID,TrueName,Tb_ClassID,LastLoginTime,UserClassName
    Rem -------------------------------------------------------------中英文截取字符串
    Public Function  cutString(str,strlen)
	If  str="" Or IsNull(str) Then 
		cutString=""
		Exit Function 
	End If 
	Dim  l,t,c, i
	str=Replace(Replace(Replace(Replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=Len(str)
	t=0
	For  i=1 To  l
		c=Abs(Asc(Mid(str,i,1)))
		If  c>255 Then 
			t=t+2
		Else 
			t=t+1
		End  If 
		If  t>=strlen Then 
			cutString=Left(str,i) & "…"
			Exit  For 
		Else 
			cutString=str
		End  If 
	Next 
	cutString=Replace(Replace(Replace(Replace(cutString," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")

	End Function

   '中英文截取字符串结束

    Rem --------------------------------------------------------字符串过滤
    Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Trim(Str)
		Str = Replace(Str,Chr(0),"")
		Str = Replace(Str,";","")
		CheckStr = Replace(Str,"'","")
	End Function
   '字符串过滤结束------

   Rem------------------------------------------------ 执行Sql语句函数
   Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDataBase()
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				Err.Clear
				Set Conn = Nothing
				Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				Response.End
			End If
		Else
			'Response.Write Command
			Set Execute = Conn.Execute(Command)
		End If
		SqlQueryNum = SqlQueryNum+1
		'Response.Write SqlQueryNum
	End Function
	'执行Sql语句函数结束----------

	Rem-------------------------清除HTML格式函数
    Public Function ClearHtml(Content)
    Content=Zxj_ReplaceHtml("&#[^>]*;",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?marquee[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?object[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?param[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?embed[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?table[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("&nbsp;","",Content)   
    Content=Zxj_ReplaceHtml("</?tr[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?th[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?p[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?a[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?img[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?tbody[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?li[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?span[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?div[^>]*>","",Content)   
    Content=Zxj_ReplaceHtml("</?th[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?td[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("</?script[^>]*>",   "",   Content)   
    Content=Zxj_ReplaceHtml("(javascript|jscript|vbscript|vbs):",   "",   Content)   
    Content=Zxj_ReplaceHtml("on(mouse|exit|error|click|key)",   "",   Content)   
    Content=Zxj_ReplaceHtml("<\\?xml[^>]*>",   "",   Content)
    Content=Zxj_ReplaceHtml("<\/?[a-z]+:[^>]*>",   "",   Content)
    Content=Zxj_ReplaceHtml("</?font[^>]*>",   "",   Content)
    'Content=Zxj_ReplaceHtml("</?b[^>]*>","",Content)
    Content=Zxj_ReplaceHtml("</?u[^>]*>","",Content)
    Content=Zxj_ReplaceHtml("</?i[^>]*>","",Content)
    Content=Zxj_ReplaceHtml("</?strong[^>]*>","",Content)
    ClearHtml=Content  
	End Function 
	'清除HTML格式函数结束
   Function   Zxj_ReplaceHtml(patrn,strng,content)   
     IF   IsNull(content)   Then   
     content=""   
     End   IF   
     Set   regEx   =   New   RegExp '   建立正则表达式。   
     regEx.Pattern   =   patrn '   设置模式。   
     regEx.IgnoreCase   =   true             '   设置忽略字符大小写。   
     regEx.Global   =   True '   设置全局可用性。   
      Zxj_ReplaceHtml=regEx.Replace(content,strng) '   执行正则匹配   
   End   Function
   Rem -------------------------------------------------------分页函数
   	Private Function URLStr(FieldName,FieldValue)
	Dim i
	For i=0 to Ubound(FieldName)
		URLStr=URLStr&"&"&CStr(FieldName(i))&"="&CStr(Checkstr(FieldValue(i)))
	Next
	End Function
	Rem iPageValue:每页个数,iRetCount:记录总数,iCurrentPage:当前页,FieldName:参数名,FieldValue:参数值
	Public Function PageList (iCurrentPage,iPageValue,iRetCount,FieldName,FieldValue)
		Dim Url
		Dim PageCount				'页总数
		Dim PageRoot				'页列表头
		Dim PageFoot				'页列表尾
		Dim OutStr
		Dim i						'输出字符串
		Const StepNum=3				'页码步长
		Url=URLStr(FieldName,FieldValue)
		If (iRetCount Mod iPageValue)=0 Then PageCount= iRetCount \ iPageValue Else PageCount= (iRetCount \ iPageValue)+1 End If
		If iCurrentPage-StepNum<=1 Then PageRoot=1 Else PageRoot=iCurrentPage-StepNum End If	
		If iCurrentPage+StepNum>=PageCount Then PageFoot=PageCount Else PageFoot=iCurrentPage+StepNum End If
		OutStr="<b class=""page"">总计 "&iRetCount&"</b> "

		If PageRoot=1 Then
			If iCurrentPage > 1 Then 
				OutStr=OutStr&"<a href='?Page="&iCurrentPage-1
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""上一页"" class=""page"">上一页</a> "
			End If
		Else
			OutStr=OutStr&"<a href='?Page=1"
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""首页"" class=""page"">首页</a> "
			OutStr=OutStr&"<a href='?Page="&iCurrentPage-1
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""上一页"" class=""page"">上一页</a> "
		End If

		For i=PageRoot To PageFoot
			If i=Cint(iCurrentPage) Then
				OutStr=OutStr&"<a class=""page"" href=""javascript:void(0);"" id=""on"">"&Cstr(i)&"</a> "
			Else
				OutStr=OutStr&"<a href='?Page="&Cstr(i)
				OutStr=OutStr&Url
				OutStr=OutStr&"' class=""page"">"&Cstr(i)&"</a> "
			End If
			If i=PageCount Then Exit For
		Next

		If PageFoot=PageCount Then
			If Cint(iCurrentPage) <> Cint(PageCount) Then 
				OutStr=OutStr&"<a href='?Page="&iCurrentPage+1
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""下一页"" class=""page"">下一页</a>"
			End If
		Else
			OutStr=OutStr&"<a href='?Page="&iCurrentPage+1
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""下一页"" class=""page"">下一页</a>"
			OutStr=OutStr&" <a href='?Page="&PageCount
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""尾页"" class=""page"">尾页</a>"
		End If
		PageList="<div id=""showpage"" style=""float:right;"">"&OutStr&"</div>"
	End Function
	'分页函数结束
	Rem ----------------------------------判断验证码是否正确
	Function CodeIsTrue()
	Dim CodeStr
	CodeStr=Trim(Request("Codestr"))
	If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
		CodeIsTrue=True
		Session("GetCode")=empty
	Else
		CodeIsTrue=False
		Session("GetCode")=empty
	End If
    End Function
	'-------------------------------------判断验证码结束
    '页面错误提示信息
   Public Function  BackInfo(ByVal tUrl,ByVal tt)
	Select Case tt
	Case 1
		Response.Write vbcrlf&"<table><tr><td colspan=2 align=center><br>"&Msg&"<br><br><a href='"&turl&"'>点击返回</a></td></tr></table><meta http-equiv='refresh' content='3; url="&tUrl&"'></td></tr>"
		Exit Function 
	Case 2
		Response.Write("<script>"&VbCrlf&"alert("""&Msg&""");window.setTimeout(""location.href='"&tUrl&"';"", 1);"&VbCrlf&"</script>")
		Exit Function
	Case Else
		Response.Write("<script>"&VbCrlf&"alert("""&Msg&""");window.setTimeout(""history.go(-1);"",1);"&VbCrlf&"</script>")
		Exit Function
	End Select
   End Function
   '验证网页是否过期
   Public Function  CheckManger()
   If Not IsArray(Session(SessionName & "JcsUser")) And Request.Cookies(CookiesName&"JcsUser") = "" Then
	Session.Abandon
    Response.Cookies(CookiesName&"JcsUser").Domain = JcsDomain
    Response.Cookies(CookiesName&"JcsUser") = ""
    Response.Cookies(CookiesName&"JcsUser").Expires = Now()-1 
	Response.Write("<SCRIPT>"&VbCrlf&"alert('登陆超时！');parent.location.href='/Index.Asp?action=exit';</SCRIPT>")
	Exit Function 
    End If
   GetMangerInfo()
   End Function 

   '---------------------------获取登陆用户信息 先判断session 如果session丢失再获取cookies
   Public Function GetMangerInfo()
   If Not IsArray(Session(SessionName & "JcsUser")) Then 
   UserID = Request.Cookies(CookiesName&"JcsUser")("Userid")
   UserName = Request.Cookies(CookiesName&"JcsUser")("UserName")
   TrueName = Request.Cookies(CookiesName&"JcsUser")("TrueName")
   LastLoginTime = Request.Cookies(CookiesName&"JcsUser")("LastLoginTime")
   UserClassName = Request.Cookies(CookiesName&"JcsUser")("UserClassName")
   UserClassID =  Request.Cookies(CookiesName&"JcsUser")("UserClassID")
   Else  
   UserInfo = Session(SessionName & "JcsUser")
   UserID = UserInfo(1)
   UserName = UserInfo(2)
   TrueName = UserInfo(3)
   LastLoginTime = UserInfo(4)
   UserClassName = UserInfo(5)
   UserClassID =  UserInfo(6)
   End If 
   End Function 

Rem -----------------------------记录操作日志;logtitle操作的标题
Public  Function WriteLog(ByVal logtitle,ByVal action)
Dim userid,truename,logcontent,logdate,user,userclassname,ip
   '先取出用户姓名
   '用户+操作时间now()+动作+标题
If  IsNull(logtitle) Or  IsNull(action)  Then 
   Founderr = True
   Exit Function 
   End If 
If Not Founderr Then
userid = Jcs.UserID
truename = Jcs.TrueName
userclassname = Jcs.UserClassName
logdate = Now() 
ip = getIP()
If IsNull(logtitle) Or logtitle = "" Then 
logcontent = userclassname & truename &"于" & logdate & action  & logtitle &"操作成功！" 
Else 
logcontent = userclassname & truename &"于" & logdate & action &"“" & logtitle &"”操作成功！"
End If 
SQL = Null
SQL = "Insert Into [Tb_Log] (UserID,Ip,LogType,LogContent,LogDate) Values ("&userid&",'"&ip&"','"&action&"','"&logcontent&"','"&logdate&"') "
Jcs.Execute(SQL)	
'Closedatabase()   '关闭数据库链接
'清空变量
'userid = 0
'truename = Null
'userclassname = Null
'logcontent = Null 
 
	End If 
End Function 

Rem--------------------------------获取客户端ip
Public  Function getIP() 
Dim strIPAddr 
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1) 
Else 
strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
End If 
getIP = Trim(Mid(strIPAddr, 1, 30)) 
End Function 
Rem--------------------------------报社下拉列表
Public  Function OfficeSelectlist(ByVal officeid)
Dim selectstr,Usr,i
SQL = "SELECT Tb_OfficeID,Tb_OfficeName FROM Tb_Office "
Set Rs = Jcs.Execute(SQL)
 If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   Msg = "没有报社信息！"
		Jcs.BackInfo "",0
    End If 
	Usr = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing     '释放Rs	
If officeid = "0" Then officeid = Null 
If IsNull(officeid)  Then  
    selectstr = selectstr + "<select name=""officename"">"
    selectstr = selectstr + "<option value="""" selected=""selected"">请选择报社</option>"
    For i = 0 To UBound(Usr,2)

    selectstr = selectstr + "<option value="&Usr(0,i)&" >"&Usr(1,i)&"</option>"
    Next
	selectstr = selectstr + "</select>"
Else 
 selectstr = selectstr + "<select name=""officename"">"
    For i = 0 To UBound(Usr,2)
	If (Usr(0,i) = officeid) Or (Usr(1,i) = officeid) Then 
    selectstr = selectstr + "<option value="&Usr(0,i)&" selected=""selected"" >"&Usr(1,i)&"</option>"
	Else 
  selectstr = selectstr + "<option value="&Usr(0,i)&" >"&Usr(1,i)&"</option>"
	End If 
    Next
	selectstr = selectstr + "</select>"
End If 
Usr = Null 
 OfficeSelectlist = selectstr
End Function 
Rem---------------------------------'系统分配随机密码
Public Function CreatePass()
	Dim Ran,i
	Createpass=""
	For i=1 To PassLength
		Randomize
		Ran = CInt(Rnd * 2)
		Randomize Timer()
		If Ran = 1 Then
			Ran = CInt(Rnd * 9)
			CreatePass = CreatePass & Ran
		Else
			Ran = CInt(Rnd * 25) + 97
			CreatePass =CreatePass & Chr(Ran)
		End If
	Next
	CreatePass= LCase(CreatePass)
End Function

Rem----------------------------------------读取用户权限
Public Function ViewUserLevel(ByVal userid)
Dim uid,userlevelid,levelstr,Usr,i,uleveli,Rstbcalss,sqltbclass,tbstr
uid = userid
If  IsNull(uid) Or Not IsNumeric(uid) Then 
   Founderr = True
    ViewUserLevel = "参数错误"
   Exit Function 
   End If 
SQL = "SELECT UserID,UserLevel FROM Jc_User WHERE UserID= "& uid
Set Rs = Jcs.Execute(SQL)
 If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   ViewUserLevel = "用户信息不存在"
    End If 
	Usr = Rs.GetRows(-1)
    userlevelid = Usr(1,0)
	If IsNull(userlevelid) Or userlevelid ="" Then 
    ViewUserLevel = "未分配权限"
	Else 
	levelstr = Split(userlevelid,",")
    For i = 0 To UBound(levelstr)
    uleveli = levelstr(i)
    sqltbclass = "SELECT Tb_classID,Tb_ClassName FROM Tb_Class WHERE Tb_classID= "& uleveli
    Set Rstbcalss = Jcs.Execute(sqltbclass)
    If Not Rstbcalss.Eof Then tbstr = tbstr + Rstbcalss(1)&"，"
	Next 
	If Not IsNull(tbstr) And tbstr<>"" Then ViewUserLevel = Left(tbstr,Len(tbstr)-1)   
	End If 
	Usr = Null 
	Rs.Close:Set Rs = Nothing     '释放Rs
	uid = Null 
	userlevelid = Null 
	levelstr = Null 
	Usr = Null 
	i= Null 
	uleveli = Null 
	Rstbcalss = Null 
	sqltbclass = Null 
	tbstr = Null 
End Function 
'--------------------------------------分配给用户多项报刊权限
Public Function CheckboxLevel(ByVal userid)
Dim uid,userlevelid,levelstr,Usr,i,uleveli,Rstbcalss,sqltbclass,tbstr,htmlstr,officeid,sqloffice,Rsclass
uid = userid
uleveli = False 
If  IsNull(uid) Or Not IsNumeric(uid) Then 
   Founderr = True
    CheckboxLevel = "参数错误"
   Exit Function 
   End If 
SQL = "SELECT UserID,UserLevel,Tb_OfficeId FROM Jc_User WHERE UserID= "& uid
Set Rs = Jcs.Execute(SQL)
 If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   CheckboxLevel = "用户信息不存在"
    End If 
	Usr = Rs.GetRows(-1)
    userlevelid = Usr(1,0)
	officeid = Usr(2,0)
	
	'尚未分配权限，列出该报社下所有报刊选项
	'先取出报社编号
	sqloffice = "SELECT Tb_classID,Tb_ClassName FROM Tb_Class WHERE (Tb_OfficeID = "&officeid & ")"
    Set Rsclass = Jcs.Execute(sqloffice)
    If Not Rsclass.eof Then 
	i = 1
	htmlstr = "<table width=""97%"" border=""0""><tr>"
	Do While Not Rsclass.eof 

	htmlstr = htmlstr + "<td height=""30"" width=""171"">"
	htmlstr = htmlstr + Rsclass(1)
	If Not IsNull(userlevelid) Then 
         If InStr(userlevelid,Rsclass(0))> 0 Then 
         htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&" checked=""checked"" /></td>"
		 Else 
         htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&" /></td>"
	     End If 
	Else 
	htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&"  /></td>"
	End If 
	If i Mod 4 = 0 Then htmlstr = htmlstr + "</tr><tr>"
    'ViewUserLevel = "未分配权限"
    Rsclass.movenext 
	i = i + 1
	Loop 
	htmlstr =  htmlstr + "</tr></table>"
	i = 0
    Rsclass.Close:Set Rsclass = Nothing     '释放Rsclass
    CheckboxLevel = htmlstr
	htmlstr = Null 
    officeid = Null 
	sqloffice = Null 
	End If 
	
	Usr = Null 
	Rs.Close:Set Rs = Nothing     '释放Rs
	'Rstbcalss.Close:Set Rstbcalss = Nothing     '释放Rs
	uid = Null 
	userlevelid = Null 
	levelstr = Null 
	Usr = Null 
	i= Null 
	uleveli = Null 
	sqltbclass = Null 
	tbstr = Null 
End Function 
'--------------------------------------分配给用户多项报纸权限
Public Function CheckPaperLevel(ByVal userid)
Dim uid,userlevelid,levelstr,Usr,i,uleveli,Rstbcalss,sqltbclass,tbstr,htmlstr,officeid,sqloffice,Rsclass
uid = userid
uleveli = False 
If  IsNull(uid) Or Not IsNumeric(uid) Then 
   Founderr = True
    CheckPaperLevel = "参数错误"
   Exit Function 
   End If 
SQL = "SELECT UserID,UserLevel FROM Jc_User WHERE UserID= "& uid
Set Rs = Jcs.Execute(SQL)
 If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   CheckPaperLevel = "用户信息不存在"
   Else 
	Usr = Rs.GetRows(-1)
    userlevelid = Usr(1,0)
	
	'尚未分配权限，列出本月内所有未分配的报纸选项
	'先取出报纸编号
	sqloffice = "SELECT PaperID,PaperName,Tb_classID,UserID FROM Tb_Paper WHERE (UserID = 0 or UserID ="&uid&")"
    Set Rsclass = Jcs.Execute(sqloffice)
    If  Rsclass.eof Then 
    CheckPaperLevel = "报纸已分配完毕"
	Else 
	i = 1
	htmlstr = "<table width=""97%"" border=""0""><tr>"
	Do While Not Rsclass.eof 

	htmlstr = htmlstr + "<td height=""30"" width=""171"">"
	htmlstr = htmlstr + Rsclass(1)
	If Not IsNull(userlevelid) Then 
         If InStr(userlevelid,Rsclass(0))> 0 Then 
         htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&" checked=""checked"" /></td>"
		 Else 
         htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&" /></td>"
	     End If 
	Else 
	htmlstr = htmlstr + "<input name=""userlevel"" type=""checkbox"" value="&Rsclass(0)&"  /></td>"
	End If 
	If i Mod 4 = 0 Then htmlstr = htmlstr + "</tr><tr>"
    'ViewUserLevel = "未分配权限"
    Rsclass.movenext 
	i = i + 1
	Loop 
	htmlstr =  htmlstr + "</tr></table>"
	i = 0
    Rsclass.Close:Set Rsclass = Nothing     '释放Rsclass
    CheckPaperLevel = htmlstr
	htmlstr = Null 
    officeid = Null 
	sqloffice = Null 
	End If 
	End If 
	Usr = Null 
	Rs.Close:Set Rs = Nothing     '释放Rs
	'Rstbcalss.Close:Set Rstbcalss = Nothing     '释放Rs
	uid = Null 
	userlevelid = Null 
	levelstr = Null 
	Usr = Null 
	i= Null 
	uleveli = Null 
	sqltbclass = Null 
	tbstr = Null 
End Function 
'------------------------------------------------------------------日期函数
Public Function ShowTime(TimeCode,N)
		Dim Ttt,D_year,D_month,D_day,D_hour,D_minute,D_second
		Ttt=Timecode
		If Ttt="" Or Isnull(Ttt) Then Ttt=Now()
		If Not(Isdate(Ttt)) Then
			Showtime="时间格式错误"
			Exit Function
		End If
		D_year=Year(Ttt)
		D_month=Month(Ttt)
		If Len(D_month)<2 Then D_month="0"&D_month
		D_day=Day(Ttt)
		If Len(D_day)<2 Then D_day="0"&D_day
		D_hour=Hour(Ttt)
		If Len(D_hour)<2 Then D_hour="0"&D_hour
		D_minute=Minute(Ttt)
		If Len(D_minute)<2 Then D_minute="0"&D_minute
		Select Case N
		Case 0
			ShowTime = D_month&"-"&D_day & " " &D_hour& ":"&D_minute
		Case 1	'2007年06月07日
			Showtime=D_year&"年"&D_month&"月"&D_day&"日"
		Case 2	'2007-06-07
			Showtime=D_year&"-"&D_month&"-"&D_day
		Case 3	'20070607
			Showtime=D_year&D_month&D_day
		Case 4	'09-20
			ShowTime = D_month&"-"&D_day
		Case Else
			Showtime=Ttt
		End Select
	End Function

End Class

%>