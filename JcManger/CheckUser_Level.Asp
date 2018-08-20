<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 用户管理" &">> 用户权限分配"
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
Call SaveUserLevel()
Case Else 
Call MainSub()
End Select 
End If 

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
  Jcs_Manger.GetUserInfo(id)
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
        <td width="25%" height="30" align="right">用户姓名：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.TrueName)%>
          &nbsp;&nbsp; <input type="hidden" name="truename" value="<%Response.Write(Jcs_Manger.TrueName)%>"></td>
      </tr>	
	  <tr>
        <td width="25%" height="30" align="right">报纸权限：</td>
        <td width="75%" class="highlight">     
<%=Jcs.CheckPaperLevel(id)%>         </td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="sub">
		 
		  <input type="submit" value="分配权限"  class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		 
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


'-------------分配报纸给用户
Sub  SaveUserLevel()
Dim manger,manuserid,userlevel,truename,i,paperlevel,SQLP,RsP
truename = Jcs.Checkstr(Request.Form("truename"))
userlevel = Jcs.Checkstr(Request.Form("userlevel"))

SQL = "Update   [Jc_User] Set UserLevel = '"&userlevel&"' Where UserID = "& CInt(ID)
Jcs.Execute(SQL)
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
Jcs.WriteLog "","给审读员“"& truename & "”分配报纸"
	Closedatabase()   '关闭数据库链接
Msg = "审读员“" & truename & "”报纸分配成功！"
		Jcs.BackInfo url & "?page="& page,2
End Sub 

%>