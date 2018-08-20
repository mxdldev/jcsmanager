<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 日志管理" &">> 日志详细信息"
'----------判断状态
Dim action
action = Request("ID")
If IsNull(action) Or action = "" Then 
Msg = "参数丢失，操作失败！"
Jcs.BackInfo "",0
Else 
Call MainSub()
End If 

'-------------------------日志显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>日志详细信息</title>
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
 Call Manageckeck()
 Jcs_Manger.GetLog(action)
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
        <td width="25%" height="30" align="right">日志操作人：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.LTrueName)%></td>
      </tr>	  
	   <tr>
        <td width="25%" height="30" align="right">操作Ip：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.Ip)%></td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">日志内容：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.LogContent)%></td>
      </tr>	
      <tr>
	    <td height="30" align="right">日志时间：	      </div></td>
        <td class="category"><%Response.Write(Jcs_Manger.LogDate)%></td>
      </tr>	     
	   <tr>
	    <td height="30" align="right">操作动作：	      </div></td>
        <td class="category"><%Response.Write(Jcs_Manger.LogType)%></td>
      </tr>	  
	  <tr>
	    <td height="30" align="right">用户身份：	      </div></td>
        <td class="category"><%= CheckUserClass()%></td>
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
Jcs_Manger.ClearLog()
End Sub
'---------------------检查管理员权限
Sub Manageckeck()
If Jcs.UserClassID <> 1 Then 
Response.Write("<br><center><img src=""../skin/images/note.gif"" align=""absmiddle"">&nbsp;<font color=""#FF0000"">你不具备此权限，请与管理员联系！</font></center>")
Response.end
End If
End Sub 
'--------------------判断身份
Function  CheckUserClass()
If Not  IsNull(Jcs_Manger.LUserClassID)  Then 
Select Case Jcs_Manger.LUserClassID
Case 1
CheckUserClass = "管理员"
Case 2
CheckUserClass = "审读员"
Case Else 
CheckUserClass = "报社用户"
End Select 
End If 
End Function 

%>
