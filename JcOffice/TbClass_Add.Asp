<!-- #include file="Const.Asp" -->
<% 

'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报刊管理" &">> 添加报刊"
'----------判断状态
Dim action
action = LCase(Request.Form("actiont"))
Select Case action
Case "sub"
Call IsExitPress()
Case Else 
Call MainSub()
End Select 
'-------------------------添加报刊主窗体部分开始
Sub MainSub()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加报刊</title>
<link href="../Skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</head>

<body>

<%
'报社用户权限
 Jcs_Office.Officeckeck()
%>

<script language="javascript">
function check()
{
if (document.form1.officename.value=="")

if (document.form1.pressname.value=="")
{
alert("报刊名称没有填写！");
document.form1.pressname.focus();
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
        <td width="75%" class="category">    <%=  Jcs_Office.GetOfficename(Jcs.UserID)%>
          &nbsp;&nbsp;</td>
      </tr>	
	    <tr>
        <td width="25%" height="30" align="right">报刊名称：</td>
        <td width="75%" class="category"><input type="text" name="pressname" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	 
	  <tr>
        <td width="25%" height="30" align="right">报刊刊号：</td>
        <td width="75%" class="category"><input type="text" name="tbnumber" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	 
	   <tr>
        <td width="25%" height="30" align="right">邮发代码：</td>
        <td width="75%" class="category"><input type="text" name="tbcode" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	 
	   <tr>
        <td width="25%" height="30" align="right">总编：</td>
        <td width="75%" class="category"><input type="text" name="editor" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	 
	   <tr>
        <td width="25%" height="30" align="right">发行周期：</td>
        <td width="75%" class="category"><input type="text" name="tbdate" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	 
	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="sub">
		  <input type="submit" value=" 确认添加 " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
		 
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

'------------判断报刊名是否存在
Sub IsExitPress()
Dim classname,tbnumber,tbcode,editor,tbdate,officeid,manuserid
manuserid = Jcs.UserID
classname = Jcs.Checkstr(Request.Form("pressname"))
officeid = Jcs_Office.GetOfficeid(Jcs.UserID)
tbnumber = Jcs.Checkstr(Request.Form("tbnumber"))
tbcode = Jcs.Checkstr(Request.Form("tbcode"))
editor = Jcs.Checkstr(Request.Form("editor"))
tbdate = Jcs.Checkstr(Request.Form("tbdate"))
If classname = ""  Then
	Founderr = True
	Msg = "报刊名称不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If

If Not Founderr Then
	SQL = "Select Tb_ClassName From [Tb_Class] Where Tb_ClassName = '"&classname&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Founderr = True
	Rs.Close:Set Rs = Nothing     '释放Rs
		Msg = "报刊名字已经存在！"
		Jcs.BackInfo "",0
	End If 
End If 

If Not Founderr Then

manuserid = Jcs.UserID
SQL = "Insert Into  [Tb_Class] (Tb_ClassName,Tb_OfficeID,Tb_Number,Tb_code,Editor,Tb_Date,UserID) Values ('"&classname&"',"&officeid&",'"&tbnumber&"','"&tbcode&"','"&editor&"','"&tbdate&"',"&manuserid&")"
Jcs.Execute(SQL)
'---写操作日志
Jcs.WriteLog classname,"添加报刊"
	Closedatabase()   '关闭数据库链接
Msg = "报刊“" & classname & "”添加成功！"
		Jcs.BackInfo "Press_List.Asp",2
End If 
End Sub 


%>
