<!-- #include file="../inc/Const.Asp" -->
<!-- #include file="../Inc/MD5.Asp" -->
<% Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'----------定义导航
Jcs.CheckManger()
Dim Navigation
Navigation = SystemName & ">> 账户管理" &">> 修改密码"
Dim action
action = LCase(Request("action"))
Select case action
case "edit"
Call EditPassWord()
case else
Call MainSub()
End Select
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>修改密码</title>
<link href="../Skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</head>

<body>



<script language="javascript">
function check()
{
if (document.form1.oldpwd.value=="")
{
alert("请输入原始密码！");
document.form1.oldpwd.focus();
return false;
}
if (document.form1.newpwd1.value=="")
{
alert("请输入新密码！");
document.form1.newpwd1.focus();
return false;
}
if (document.form1.newpwd2.value=="")
{
alert("请在此输入新密码！");
document.form1.newpwd2.focus();
return false;
}
if(document.form1.newpwd2.value!=document.form1.newpwd1.value)
{
alert("两次输入密码不一致！");
document.form1.newpwd2.focus();
return false;
}
}
</script>
<form name="form1" ACTION="" >
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
	  <tr>
        <td width="25%" height="30" align="right">用户名：</td>
        <td width="75%" class="category"><%=Jcs.UserName%>          &nbsp;&nbsp;</td>
      </tr>	 
	   <tr>
        <td width="25%" height="30" align="right">原始密码：</td>
        <td width="75%" class="category">
          <input type="password" name="oldpwd" value="" style="width:300px"></td>
      </tr>	
	  <tr>
        <td width="25%" height="30" align="right">新密码：</td>
        <td width="75%" class="category">
          <input type="password" name="newpwd1" value="" style="width:300px"></td>
      </tr>	
	  <tr>
        <td width="25%" height="30" align="right">确认新密码：</td>
        <td width="75%" class="category">
          <input type="password" name="newpwd2" value="" style="width:300px"></td>
      </tr>	
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="action" value="edit">
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

'-------------添加保存公告信息
Sub  EditPassWord()
Dim oldpwd,newpwd1,newpwd2
Dim Usr,Pw,str,sql1,str1
oldpwd=Jcs.Checkstr(Request("oldpwd"))
newpwd1=Jcs.Checkstr(Request("newpwd1"))
newpwd2=Jcs.Checkstr(Request("newpwd2"))
SQL = "Select UserID,UserName,UserPassWord From Jc_User Where UserId="&Jcs.UserId

	Set Rs = Jcs.Execute(SQL)
	If Not Rs.Bof Then
    Usr = Rs.GetRows()
	Rs.Close:Set Rs = Nothing     '释放Rs
	
	str = Left(Usr(2,0),PassLength)
	PW = str & MD5(str & oldpwd)
	If PW <> Usr(2,0) Then
		Founderr = True
		Msg = "原始密码输入有误！"
		Jcs.BackInfo "",0
	Else
	    str1=Jcs.CreatePass
		sql1="Update Jc_User Set UserPassWord='"&str1&MD5(str1&newpwd1)&"' Where UserId="&Jcs.UserId
	    Jcs.Execute(sql1)
		Jcs.WriteLog "修改密码","修改密码"
	    Closedatabase()   '关闭数据库链接
       Msg = "密码修改成功！"
		Jcs.BackInfo "PassWordEdit.Asp",2
	End If 
	
	End If
'---写操作日志

End Sub 

%>