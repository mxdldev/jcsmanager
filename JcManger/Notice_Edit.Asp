<!-- #include file="Const.Asp" -->
<% Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 公告管理" &">> 公告名称修改"
'----------判断状态
Dim id,action
id = Request("ID")
If IsNull(id) Or id = "" Then 
Msg = "参数丢失，操作失败！"
Jcs.BackInfo "",0
Else 
action = Request("actiont")
Select Case action
Case "sub"
Call IsExitNotice()
Case Else 
Call MainSub()
End Select 
End If 

'-------------------------公告显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>公告名称修改</title>
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
  Jcs_Manger.GetNotice(id)
%>

<script language="javascript">
function check()
{
if (document.form1.Noticename.value=="")
{
alert("公告姓名没有填写！");
document.form1.Noticename.focus();
return false;
}
}
</script>
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
        <td width="25%" height="30" align="right">公告名称：</td>
        <td width="75%" class="category"><input type="text" name="Noticename" value="<%Response.Write(Jcs_Manger.NoticeTitle)%>" style="width:300px">
          &nbsp;&nbsp;</td>
      </tr>	 
	   <tr>
        <td width="25%" height="30" align="right">公告内容：</td>
        <td width="75%" class="category">
          <textarea name="noticecontent" style="display:none"><%=Response.Write(Jcs_Manger.NoticeContent) %></textarea>
             <iframe ID="Editor" name="Editor" src="../JcsEditor/index.html?ID=noticecontent" frameBorder="0" marginHeight="0" marginWidth="0" scrolling="No" style="height:248px;width:550px"></iframe></td>
      </tr>	
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

'------------判断公告名是否存在
Sub IsExitNotice()
Dim Noticename,noticecontent
Noticename = Jcs.Checkstr(Request.Form("Noticename"))
If Noticename = ""  Then
	Founderr = True
	Msg = "公告名称不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If
noticecontent = Jcs.Checkstr(Request.Form("noticecontent"))
If noticecontent = ""  Then
	Founderr = True
	Msg = "公告内容不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If

If Not Founderr Then
Call EditSaveNotice()
End If 
End Sub 

'-------------添加保存公告信息
Sub  EditSaveNotice()
SQL = "Update   [Tb_Notice] Set NoticeTitle = '"&Jcs.Checkstr(Request.Form("Noticename"))&"' , NoticeContent = '"&Jcs.Checkstr(Request.Form("noticecontent"))&"' Where NID = "& id
Jcs.Execute(SQL)
'---写操作日志
Jcs.WriteLog Jcs.Checkstr(Request.Form("Noticename")),"修改公告"
	Closedatabase()   '关闭数据库链接
Msg = "公告修改成功！"
		Jcs.BackInfo "Notice_List.Asp",2
End Sub 

%>