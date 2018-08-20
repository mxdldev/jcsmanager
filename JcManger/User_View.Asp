<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报社用户管理" &">> 用户详细信息"
'----------判断状态
Dim action,url,page
action = Request("ID")
url = Request("url")
page = Request("Page")
If IsNull(action) Or action = "" Then 
Msg = "参数丢失，操作失败！"
Jcs.BackInfo "",0
Else 
Call MainSub()
End If 

'-------------------------用户显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户详细信息</title>
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
 'Jcs_Manger.Manageckeck()
 Jcs_Manger.GetUserInfo(action)
%>

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
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.UserName)%></td>
      </tr>	  
	   <tr>
        <td width="25%" height="30" align="right">真实姓名：</td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.TrueName)%></td>
      </tr>	
	  
      <tr>
	    <td height="30" align="right">锁定状态	      </div></td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.cndeleteflag)%></td>
      </tr>	 
	   <%If Jcs_Manger.UserClassName <> "审读员"   Then 
	  Response.write("")
	  Else %> 
	   <tr>
        <td width="25%" height="30" align="right">所属报社：</td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.UofficeName)%></td>
      </tr>	
	   <tr>
	    <td height="30" align="right">权限：	      </div></td>
        <td width="63%" class="category"><%Response.Write(Jcs.ViewUserLevel(action))%></td>
        <td width="12%" ><a href="CheckUser_Level.Asp?id=<%=action%>&url=<%=url%>&page=<%=page%>">分配权限</a></td>
	   </tr>	
	   <% End If %>
	   <tr>
	    <td height="30" align="right">用户类别：	      </div></td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.UserClassName)%></td>
      </tr>	    
	   <tr>
	    <td height="30" align="right">最后一次登陆：	      </div></td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.LastLoginTime)%></td>
      </tr>	 
	   <tr>
	    <td height="30" align="right">登陆次数：	      </div></td>
        <td colspan="2" class="category"><%Response.Write(Jcs_Manger.Logintimes)%></td>
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

</body>
</html>
<% '清空变量
Jcs_Manger.ClearUserInfo()
End Sub
'----------检查用户权限
Function CheckUserLevl(id)
'根据权限编号获取报刊名     

'先取出权限编号集合
Dim levelstr,leveli,i,levelhtstr
levelstr = Jcs_Manger.UserLevel
leveli = Split(levelstr,",")
For i = 0 To UBound(leveli)
'执行sql语句找到报刊 

levelhtstr = levelhtstr + "，"
Next 
CheckUserLevl = levelhtstr
End Function 
%>
