<!-- #include file="Const.Asp" -->
<script language="javascript" type="text/javascript" src="../skin/Js/Jcs_Function.js"></script>
<% 
'判断登陆
Jcs.CheckManger()
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报社管理" &">> 添加报社"
'----------判断状态
Dim action
action = LCase(Request.Form("actiont"))
Select Case action
Case "sub"
Call IsExitOffice()
Case Else 
Call MainSub()
End Select 
'-------------------------报社添加主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报社添加</title>
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

<script language="javascript">
function check()
{
if (document.form1.Officename.value=="")
{
alert("报社姓名没有填写！");
document.form1.Officename.focus();
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
        <td width="25%" height="30" align="right">报社名称：</td>
        <td width="75%" class="category"><input type="text" name="Officename" onKeyUp="javascript:document.getElementById('Submit').disabled = true;"  Onblur="ajaxLoadPage('ajaxdata.asp?action=office&ID=',this.value,'get','IsOkUserName','1.1.3');" style="width:250px">
          &nbsp;&nbsp;<span id="IsOkUserName"></span></td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="sub">
		  <input type="submit" value=" 确认添加 " id="Submit" onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		 
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
<script language="javascript" defer="defer">
document.getElementById('Submit').disabled = true;
</script>
</body>
</html>
<%
End Sub

'------------判断报社名是否存在
Sub IsExitOffice()
Dim officename
officename = Jcs.Checkstr(Request.Form("Officename"))
If officename = ""  Then
	Founderr = True
	Msg = "报社名称不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If

If Not Founderr Then
Call addSaveoffice()
End If 
End Sub 

'-------------添加保存报社信息
Sub  addSaveoffice()
Dim manger,manuserid
manger = Jcs.UserName
manuserid = Jcs.UserID
SQL = "Insert Into  [Tb_Office] (Tb_OfficeName,Manager,UserID) Values ('"&Jcs.Checkstr(Request.Form("Officename"))&"','"&manger&"',"&manuserid&")"
Jcs.Execute(SQL)
'---写操作日志
Jcs.WriteLog Jcs.Checkstr(Request.Form("Officename")),"添加报社"
	Closedatabase()   '关闭数据库链接
Msg = "报社添加成功！"
		Jcs.BackInfo "Office_List.Asp",2
End Sub 

%>
