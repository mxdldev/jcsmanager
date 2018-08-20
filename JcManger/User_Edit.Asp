<!-- #include file="Const.Asp" -->
<!-- #include file="../Inc/MD5.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 用户管理" &">> 用户信息修改"
'----------判断状态
Dim id,action,url,page
id = Request("ID")
If IsNull(id) Or id = "" Then 
Msg = "参数丢失，操作失败！"
Jcs.BackInfo "",0
Else 
action = Request("actiont")
url = Request("url")
page = Request("Page")
Select Case action
Case "sub"
Call EditSaveUser()
Case Else 
Call MainSub()
End Select 
End If 

'-------------------------用户显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户信息修改</title>
<link href="../Skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</head>

<body>

<%

'管理员权限
 ' Jcs_Manger.Manageckeck()
  Jcs_Manger.GetUserInfo(id)
%>

<script language="javascript">
  function check()
{
var patrn=/^[a-zA-Z][a-zA-Z0-9_]{4,15}$/;
var parpwdmac=/^(\w){6,20}$/;
if (document.formadd.officename.value=="")
{
alert("请选择报社！");
document.formadd.officename.focus();
return false;
}
if (document.formadd.username.value=="")
{
alert("用户名称没有填写！");
document.formadd.username.focus();
return false;
}
if (!patrn.exec(document.formadd.username.value))
{
alert("用户名格式不正确！");
document.formadd.username.focus();
return false;
}

if (document.formadd.repassword.value!=document.formadd.password.value)
{
alert("两次密码不一样，请重新输入！");
document.formadd.repassword.focus();
return false;
}
if (document.formadd.truename.value=="")
{
alert("真实姓名没有填写！");
document.formadd.truename.focus();
return false;
}
}
//
</script>
<form name="formadd" ACTION="" METHOD="post" >
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../skin/images/r_1.gif" alt="" /></td>
<td width="100%" background="../skin/images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;<%= Navigation%></td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../skin/images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
	  <%If Jcs_Manger.UserClassName <> "报社用户" Then 
	  Response.write("")
	  Else %> 
		<tr>
        <td width="25%" height="30" align="right">所属报社：</td>
        <td width="75%" class="category">
          &nbsp;&nbsp;<%= Jcs.OfficeSelectlist(Jcs_Manger.UofficeName)%></td>
      </tr>
	  <%End If %>
	 <tr>
        <td width="25%" height="30" align="right">用户名：</td>
        <td width="75%" class="category"><input type="text" name="username" value="<%Response.Write(Jcs_Manger.UserName)%>" style="width:200px">
          &nbsp;&nbsp;(字母开头，允许5-16字节，允许字母数字下划线)</td>
      </tr>	
	  
	   <tr>
        <td width="25%" height="30" align="right">新密码：</td>
        <td width="75%" class="category"><input  type="password" name="password" style="width:200px">
          &nbsp;&nbsp;(为空默认的原密码)</td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">重新输入密码：</td>
        <td width="75%" class="category"><input  type="password" name="repassword" style="width:200px">
          &nbsp;&nbsp;(6-20个字母、数字、下划线)</td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">真实姓名：</td>
        <td width="75%" class="category"><input type="text" name="truename" value="<%Response.Write(Jcs_Manger.TrueName)%>" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>
	   <%If Jcs_Manger.UserClassName <> "审读员" Then 
	  Response.write("")
	  Else %> 
		<tr>
        <td width="25%" height="30" align="right">报纸权限：</td>
        <td width="75%" class="highlight">     
<%=Jcs.CheckPaperLevel(id)%> 
         </td>
      </tr>	  
	  <%End If %>
	   
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="sub">
		  <input type="submit" value=" 确认修改 " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		 
		  <input type="reset" value=" 重新填写 " class="button">		</td>
      </tr>	    
</table>
</td>
<td></td>
</tr>
<tr>
<td><img src="../skin/images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../skin/images/r_3.gif" alt="" /></td>
</tr>
</table>
</form>
</body>
</html>
<%
End Sub

'------------保存用户信息
Sub EditSaveUser()
Dim truename,username,pwd,Tbofficeid,ND,userlevel,i,paperlevel,SQLP,RsP
truename = Jcs.Checkstr(Request.Form("truename"))
username = Jcs.Checkstr(Request.Form("username"))
pwd = Jcs.Checkstr(Request.Form("password"))

Tbofficeid = Jcs.Checkstr(Request.Form("officename"))
If Tbofficeid ="" Or IsNull(Tbofficeid) Then Tbofficeid = 0
userlevel = Jcs.Checkstr(Request.Form("userlevel"))
If truename = ""  Then
	Founderr = True
	Msg = "真实姓名不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If
If username = ""  Then
	Founderr = True
	Msg = "用户名不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If

If Not Founderr Then
If pwd = "" Or IsNull(pwd) Then 

SQL = "Update   [Jc_User] Set UserName='"&username&"',TrueName='"&truename&"',Tb_OfficeId="&Tbofficeid&",UserLevel = '"&userlevel&"'  Where UserID = "& id
Else
ND = Jcs.CreatePass
Pwd= ND & MD5(ND & pwd)
SQL = "Update   [Jc_User] Set UserName='"&username&"',TrueName='"&truename&"',UserPassWord='"&pwd&"',Tb_OfficeId="&Tbofficeid&",UserLevel = '"&userlevel&"'  Where UserID = "& id
End If 
Jcs.Execute(SQL)
'

'--------------写入报纸的审读员编号
'paperlevel = Split(userlevel,",")
SQLP = "SELECT PaperID, PaperName, UserID, Tb_Check FROM Tb_Paper WHERE (Tb_Check = 0) AND (UserID = "&CInt(ID)&" OR UserID = 0)"
Set RsP = Jcs.Execute(SQLP)
Do While Not RsP.eof 
If InStr(userlevel,RsP(0))> 0 Then 
SQL = "Update   [Tb_Paper] Set UserID = "&CInt(ID)&" Where PaperID = "& CInt(RsP(0))
Else 
SQL = "Update   [Tb_Paper] Set UserID = 0 Where  PaperID = "& CInt(RsP(0))
End If 
Jcs.Execute(SQL)
RsP.movenext 
Loop 
'---写操作日志
Jcs.WriteLog Jcs.Checkstr(Request.Form("truename")),"修改用户信息"
	Closedatabase()   '关闭数据库链接
Msg = "用户信息修改成功！"
		Jcs.BackInfo url & "?page="& page,2
End If 
End Sub 

%>