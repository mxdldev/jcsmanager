<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 公告管理" &">> 公告详细信息"
'----------判断状态
Dim action
action = Request("ID")
If IsNull(action) Or action = "" Then 
Msg = "参数丢失，操作失败！"
Jcs.BackInfo "",0
Else 
Call MainSub()
End If 

'-------------------------公告显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>公告详细信息</title>
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
 Jcs_Manger.GetNotice(action)
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
        <td width="25%" height="30" align="right">公告名称：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.NoticeTitle)%></td>
      </tr>	  
	   <tr>
        <td width="25%" height="30" align="right">公告内容：</td>
        <td width="75%" class="category"><%Response.Write(Jcs_Manger.NoticeContent)%></td>
      </tr>	
	   
	   <tr>
	    <td height="30" align="right">公告添加时间：	      </div></td>
        <td class="category"><%Response.Write(Jcs_Manger.NAddDate)%></td>
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
Jcs_Manger.ClearNotice()
End Sub


%>
