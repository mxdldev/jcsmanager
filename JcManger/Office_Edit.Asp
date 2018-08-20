<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报社管理" &">> 报社名称修改"
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
Call IsExitOffice()
Case Else 
Call MainSub()
End Select 
End If 

'-------------------------报社显示主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报社名称修改</title>
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
  Jcs_Manger.GetOffice(id)
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
        <td width="75%" class="category"><input type="text" name="Officename" value="<%Response.Write(Jcs_Manger.Tb_OfficeName)%>" style="width:300px">
          &nbsp;&nbsp;</td>
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
	SQL = "Select Tb_OfficeName From [Tb_Office] Where Tb_OfficeName = '"&officename&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Founderr = True
	Rs.Close:Set Rs = Nothing     '释放Rs
		Msg = "报社名字已经存在！"
		Jcs.BackInfo "",0
	End If 
End If 

If Not Founderr Then
Call EditSaveoffice()
End If 
End Sub 

'-------------添加保存报社信息
Sub  EditSaveoffice()
Dim manger,manuserid
manger = Jcs.UserName
manuserid = Jcs.UserID
SQL = "Update   [Tb_Office] Set Tb_OfficeName = '"&Jcs.Checkstr(Request.Form("Officename"))&"' Where Tb_OfficeID = "& id
Jcs.Execute(SQL)
'---写操作日志
Jcs.WriteLog Jcs.Checkstr(Request.Form("Officename")),"修改报社名称"
	Closedatabase()   '关闭数据库链接
Msg = "报社名称修改成功！"
		Jcs.BackInfo "Office_List.Asp",2
End Sub 
'---------------获取报社用编号
Function GetOfficeid(userid)
Dim Rs
If IsNull(userid) Then 
GetOfficeid = ""
Exit Function 
End If 
SQL = "SELECT Tb_OfficeId FROM Jc_User WHERE (UserID = "&userid&")"
Set Rs = Jcs.Execute(SQL)
If Not rs.eof Then GetOfficeid = Rs(0)
Rs.Close:Set Rs = Nothing     '释放Rs
End Function 

%>