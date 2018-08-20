<!-- #include file="Inc/Const.Asp" -->
<!-- #include file="Inc/MD5.Asp" -->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'----------判断是登陆、验证用户登陆或退出
Dim action
action = LCase(Request("action"))
Select Case action
Case "sub"
Call CheckLogin()
Case "exit"
Call LogJsout()
Case Else 
Call MainSub()
End Select 
'-------------------------系统登陆主窗体部分开始
Sub MainSub()
Dim Htmlstr 
Htmlstr = Htmlstr + "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
Htmlstr = Htmlstr + "<html xmlns=""http://www.w3.org/1999/xhtml"">"
Htmlstr = Htmlstr + "<head>"
Htmlstr = Htmlstr + "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"
Htmlstr = Htmlstr + "<title>" & JccTitleName & " - 登陆入口</title>"
Htmlstr = Htmlstr + "<STYLE>"
Htmlstr = Htmlstr + "body {font-family: ""宋体"";font-size: 12px;text-decoration: none;}"
Htmlstr = Htmlstr + "td {font-size: 12px;color: #666666;text-decoration: none;line-height: 18px;font-family: ""宋体"";}"
Htmlstr = Htmlstr + ".S1{"
Htmlstr = Htmlstr + "font-family: ""宋体"";"
Htmlstr = Htmlstr + "font-weight: bold;"
Htmlstr = Htmlstr + "color: #ffffff;"
Htmlstr = Htmlstr + "font-size:16px;"
Htmlstr = Htmlstr + "text-decoration: none}"
Htmlstr = Htmlstr + "</STYLE>"
Htmlstr = Htmlstr + "<SCRIPT language=JavaScript type=text/JavaScript>"
Htmlstr = Htmlstr + "nereidFadeObjects = new Object();"
Htmlstr = Htmlstr + "nereidFadeTimers = new Object();"
Htmlstr = Htmlstr + "function nereidFade(object, destOp, rate, delta){"
Htmlstr = Htmlstr + "if (!document.all)"
Htmlstr = Htmlstr + "return;"
Htmlstr = Htmlstr + "if (object != ""[object]""){"
Htmlstr = Htmlstr + "setTimeout(""nereidFade"
Htmlstr = Htmlstr + "(""+object+"",""+destOp+"",""+rate+"",""+delta+"")"",0);"
Htmlstr = Htmlstr + "return;"
Htmlstr = Htmlstr + "return;"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "clearTimeout(nereidFadeTimers[object.sourceIndex]);"
Htmlstr = Htmlstr + "diff = destOp-object.filters.alpha.opacity;"
Htmlstr = Htmlstr + "direction = 1;"
Htmlstr = Htmlstr + "if (object.filters.alpha.opacity > destOp){"
Htmlstr = Htmlstr + "direction = -1;"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "delta=Math.min(direction*diff,delta);"
Htmlstr = Htmlstr + "object.filters.alpha.opacity+=direction*delta;"
Htmlstr = Htmlstr + "if (object.filters.alpha.opacity != destOp){"
Htmlstr = Htmlstr + "nereidFadeObjects[object.sourceIndex]=object;"
Htmlstr = Htmlstr + "nereidFadeTimers[object.sourceIndex]=setTimeout(""nereidFade"
Htmlstr = Htmlstr + "(nereidFadeObjects[""+object.sourceIndex+""],""+destOp+"",""+rate+"",""+delta+"")"",rate);"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "function document.onreadystatechange()"
Htmlstr = Htmlstr + "{  var app=navigator.appName;"
Htmlstr = Htmlstr + "var verstr=navigator.appVersion;"
Htmlstr = Htmlstr + "if(app.indexOf(""Netscape"") != -1) {"
Htmlstr = Htmlstr + "alert(""友情提示：\n    您使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。"");"
Htmlstr = Htmlstr + " } else if(app.indexOf(""Microsoft"") != -1) {"
Htmlstr = Htmlstr + "if (verstr.indexOf(""MSIE 3.0"")!=-1 || verstr.indexOf(""MSIE 4.0"") != -1 || verstr.indexOf(""MSIE 5.0"") != -1 || verstr.indexOf(""MSIE 5.1"") != -1)"
Htmlstr = Htmlstr + "alert(""友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。"");"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + " document.loginform.username.focus();"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "function checkform() {"
Htmlstr = Htmlstr + "if(loginform.username.value == """") {"
Htmlstr = Htmlstr + "alert(""请输入账号！"");"
Htmlstr = Htmlstr + "loginform.username.focus();"
Htmlstr = Htmlstr + "return false;"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "if(loginform.pwd.value == """") {"
Htmlstr = Htmlstr + "alert(""请输入密码！"");"
Htmlstr = Htmlstr + "loginform.pwd.focus();"
Htmlstr = Htmlstr + "return false;"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "if (loginform.verifycode.value == """") {"
Htmlstr = Htmlstr + "alert (""请输入验证码！"");"
Htmlstr = Htmlstr + "loginform.verifycode.focus();"
Htmlstr = Htmlstr + "return false;"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "}"
Htmlstr = Htmlstr + "</SCRIPT>"
Htmlstr = Htmlstr + "</head>"
Htmlstr = Htmlstr + "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
Htmlstr = Htmlstr + "<table width=""95%"" border=""0"">"
Htmlstr = Htmlstr + "<tr>"
Htmlstr = Htmlstr + " <td height=""600"" ><table width=""646"" height=""425"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
Htmlstr = Htmlstr + "<tr><td colspan=""3""><img src=""skin/Images/login_01.gif"" width=""646"" height=""114"" alt=""""></td></tr>"
Htmlstr = Htmlstr + "<tr><td><img src=""skin/Images/login_02.gif"" width=""88"" height=""311"" alt=""""></td><td width=""476"" height=""311"" valign=""top"" background=""skin/images/login_03.gif""><br>"
Htmlstr = Htmlstr + "<FORM ACTION="""" METHOD=""post"" name=""loginform"" ><table width=""80%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
Htmlstr = Htmlstr + "<tr><td width=""87"" height=""30"" align=""right"">登录名：</td><td valign=""middle""><input type=""text"" name=""username"" tabindex=""1"" style=""width:160px;""></td></tr>"
Htmlstr = Htmlstr + "<tr><td height=""30"" align=""right"">密　码：</td><td><INPUT name=""pwd""  type=""password"" tabindex=""2"" style=""width:160px;""></tr>"
Htmlstr = Htmlstr + "<tr><td height=""30"" align=""right"">验证码：</td><td height=""30""><input type=""text"" class=""vcode"" maxlength=4 id=""verifycode"" tabindex=""4"" name=""codestr"" size=""6""> <font color=red><-请在验证码框输入</font><img src=""Skin/images/GetCode.asp"" id=""checkcode"" align=""absmiddle""> <A HREF=""###"" onclick=""checkcode.src=checkcode.src"">更换一个？</A>&nbsp;</td></tr>"
Htmlstr = Htmlstr + "<tr><td height=""50"" colspan=""2"" align=""center""><input onMouseOver=nereidFade(this,100,10,5)  style=""FILTER:alpha(opacity=50)"" onMouseOut=nereidFade(this,50,10,5) src=""skin/Images/dl.gif""  type=""image"" tabindex=""5"" name=""enter"" border=""0"" onClick=""return checkform()""></td><input type=""hidden"" name=""action"" value=""sub"" /></td></tr>"
Htmlstr = Htmlstr + "<tr><td colspan=""2""></td></tr></table></form><hr size=1 color=""#efefef"" width=""80%"" align=""left"">"
Htmlstr = Htmlstr + "<div align=""center"">Copyright &copy 2009-2010 Powered By  锦华科技 Version 1.1<br></div>"
Htmlstr = Htmlstr + "</td><td><img src=""skin/Images/login_04.gif"" width=""82"" height=""311"" alt=""""></td></tr></table></td></tr></table></body></html>"
Response.Write(Htmlstr)
%>

<%
End Sub 
'-------------------主窗体登陆部分结束

'------CheckLogin()用户登陆验证

Sub CheckLogin()
Dim name,Password,codestr,Founderr
Dim Usr,Pw,str,N,i
Dim Key,iMyUser
name = Jcs.Checkstr(Request.Form("username"))
Password = Jcs.Checkstr(Request.Form("pwd"))
codestr = Jcs.Checkstr(Request.Form("codestr"))
If name = "" Or Password = "" Then
	Founderr = True
	Msg = "用户名或者密码错误！"
	Jcs.BackInfo "",0
	Exit Sub
End If
If Not Jcs.CodeIsTrue() Then
	Founderr = True
	Msg = "错误的验证码！"
	Jcs.BackInfo "",0
	Exit Sub
End If
If Not Founderr Then
	SQL = "Select UserID,UserName,UserPassWord,UserClassID,TrueName,Tb_ClassID,LastLoginTime,UserClassName,DeleteFlag From [v_JcUser] Where DeleteFlag <> 1 And UserName = '"&name&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Rs.Bof Or Rs.Eof Then
	Founderr = True
		Msg = "用户名或密码错误！"
		Jcs.BackInfo "",0
	Else 
    Usr = Rs.GetRows()
	Rs.Close:Set Rs = Nothing     '释放Rs
	
	str = Left(Usr(2,0),PassLength)
	PW = str & MD5(str & Password)
	If PW <> Usr(2,0) Then
		Founderr = True
		Msg = "用户名或密码错误！"
		Jcs.BackInfo "",0
	End If 
	If  Usr(8,0) = 2 Then     '判断是否被锁定
		Founderr = True
		Msg = "用户名被锁定，请与管理员联系！"
		Jcs.BackInfo "",0
	End If 
End If
End If 
If Not Founderr Then
	N = ArrSafeCode & "$$" & Usr(0,0) & "$$" & Usr(1,0) & "$$" & Usr(4,0) & "$$" & Usr(6,0) & "$$" & Usr(7,0)  & "$$" & Usr(3,0) & "$$" & ArrSafeCode
	iMyUser = Split(N,"$$")    'session数组里存放 1用户编号，2用户名，3真实姓名，4最后一次登陆时间，5用户级别，6用户等级编号
	'设置session和cookie
	Session(SessionName & "JcsUser") = iMyUser
	Response.Cookies(CookiesName&"JcsUser")("Userid") = Usr(0,0)
    Response.Cookies(CookiesName&"JcsUser")("UserName") = Usr(1,0)
	Response.Cookies(CookiesName&"JcsUser")("TrueName") = Usr(4,0)
	Response.Cookies(CookiesName&"JcsUser")("LastLoginTime") = Usr(6,0)
	Response.Cookies(CookiesName&"JcsUser")("UserClassName") = Usr(7,0)
	Response.Cookies(CookiesName&"JcsUser")("UserClassID") = Usr(3,0)
	
	SQL = "Update [Jc_User] Set Logintimes = Logintimes + 1,LastLoginTime = "&SqlNowString&" Where UserID = " & iMyUser(1)
	Jcs.Execute(SQL)
	Closedatabase()   '关闭数据库链接
	
	Response.Redirect "Main.Asp"
End If
End Sub 
'------------Logout() 退出系统
Sub LogJsout()
   Session.Abandon
   Response.Cookies(CookiesName&"JcsUser").Domain = JcsDomain
   Response.Cookies(CookiesName&"JcsUser") = ""
   Response.Cookies(CookiesName&"JcsUser").Expires = Now()-1 
   'Msg = "安全退出！"
   'Jcs.BackInfo "/Index.Asp",2
   Response.Redirect "/"
   Response.end
End Sub 

%>
