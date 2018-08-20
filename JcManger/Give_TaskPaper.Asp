<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 用户管理" &">> 用户权限分配"
'----------判断状态
Dim id,action,url,page
id = Request("ID")

action = Request("actiont")
url = Request("url")
page = Request("Page")
Select Case action
Case "sub"
Call SaveUserLevel()
Case Else 
Call MainSub()
End Select 


'-------------------------报社用户显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报社用户权限分配</title>
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
  Jcs_Manger.Manageckeck()
%>

<form name="form1" ACTION="" METHOD="post" >
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
        <td width="25%" height="30" align="right">所有未分配报纸：</td>
        <td width="75%" class="category"><%=Jcs_Manger.CheckAllPaper()%> 
          &nbsp;&nbsp; </td>
      </tr>	
	  <tr>
        <td width="25%" height="30" align="right">选择用户：</td>
        <td width="75%" class="highlight">     
<%=Jcs_Manger.ListAllcheckuser()%>       
<p>
  <label></label>
  <br />
</p></td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="sub">
		 
		  <input type="submit" value="分配任务"  onClick="return check()"  class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		 
		  <input type="reset" value=" 重新分配 " class="button">		</td>
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


'-------------分配报纸给用户
Sub  SaveUserLevel()
Dim manger,manuserid,userlevel,truename,i,paperlevel,SQLP,Rs,RsP,UserList,levelstr,paperstr,Sqllevel
manger = Jcs.TrueName
userlevel =  Jcs.Checkstr(Request.Form("userlevel"))
UserList = Jcs.Checkstr(Request.Form("UserList"))
If IsNull(userlevel) Or userlevel="" Then 
Msg = "至少选择一个报纸选项！"
		Jcs.BackInfo "",0
		Exit Sub 
End If 

If IsNull(UserList) Or UserList="" Then 
Msg = "请选择用户！"
		Jcs.BackInfo "",0
		Exit Sub 
End If 
'先取出用户的权限levelstr
SQL = "SELECT UserLevel, TrueName, UserID FROM Jc_User WHERE (UserID = "&CInt(UserList)&")"
Set Rs = Jcs.Execute(SQL)
	If  Not  Rs.Bof Then
    levelstr = Rs(0)
	truename = Rs(1)
	Rs.Close:Set Rs = Nothing     '释放Rs
    End If 
'循环拼接报纸字符串
paperstr = ""
'If  IsNull(levelstr) Or levelstr<> "" Then levelstr=levelstr&","
'Response.write(levelstr)
paperlevel = Split(userlevel,",")
For i = 0 To UBound(paperlevel)
If InStr(levelstr,paperlevel(i)) = 0 Then paperstr = paperstr +  paperlevel(i)&","
'写入报纸用户编号
SQL = "Update   [Tb_Paper] Set UserID = "&CInt(UserList)&" Where PaperID = "&CInt(paperlevel(i))
Jcs.Execute(SQL)
Next 
'Response.write(paperstr)
If  IsNull(levelstr) Or levelstr<> "" Then levelstr=levelstr&","
Sqllevel = levelstr & Left(paperstr,Len(paperstr) -1)
'Response.write(Sqllevel)
'给用户分配报纸权限
SQL = "Update   [Jc_User] Set UserLevel = '"&Sqllevel&"' Where UserID = "&CInt(UserList)
Jcs.Execute(SQL)

'写操作日志
Jcs.WriteLog "","给审读员用户“"& truename & "”分配报纸"
	Closedatabase()   '关闭数据库链接
Msg = "给审读员用户“" & truename & "”权限分配成功！"
			Jcs.BackInfo url & "?",2
levelstr = Null 
SQL = Null 
paperlevel = Null 
paperstr =Null 
userlevel =Null 
UserList =Null 
End Sub 

%>