<!-- #include file="Const.Asp" -->
<% 
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报刊审核管理" &">> 审核报刊"
'----------判断状态
Dim action,id
id= Request("ID")
action=Request("action")
Select Case action
Case "show"
Call ShowTbClass(id)
Case "review"
Call Jcs_Check.ClassReview()
Case Else
Msg = "错误的来源！"
Jcs.BackInfo "/Desk.asp",2
End Select

'-------------------------公告显示主窗体部分开始
Sub ShowTbClass(id)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报刊详细信息</title>
<link href="../Skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</head>

<body>

<%
Dim Sql,Rs,Ds
Sql="Select Tb_ClassId,Tb_ClassName,Tb_Number,Tb_Code,Editor,Tb_Date,Tb_ClassState From Tb_Class Where Tb_ClassId="&id
'Response.Write sql
'Response.End()
Set Rs=Jcs.Execute(Sql)
If Not Rs.eof Then Ds=Rs.Getrows()
If IsArray(Ds) Then
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
<form>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
	  <tr>
        <td width="25%" height="30" align="right">报刊名称：</td>
        <td width="75%" class="category"><%=Ds(1,0)%></td>
      </tr>	  
	   <tr>
        <td width="25%" height="30" align="right">刊号：</td>
        <td width="75%" class="category"><%=Ds(2,0)%></td>
      </tr>	
	   
	   <tr>
	    <td height="30" align="right">邮发代号：	      </div></td>
        <td class="category"><%=Ds(3,0)%></td>
      </tr>	   
	  <tr>
	    <td height="30" align="right">总编辑：	      </div></td>
        <td class="category"><%=Ds(4,0)%></td>
      </tr>	 
	  <tr>
	    <td height="30" align="right">发型周期：	      </div></td>
        <td class="category"><%=Ds(5,0)%></td>
      </tr>	
	    <tr>
	    <td height="30" align="right">审核：	      </div></td>
        <td class="category">通过<input type="radio" name="stat" value=1 style="border:none"
		<%
		 If Cint(Ds(6,0))=1 Then Response.Write "checked"
		%>
		>&nbsp;&nbsp;&nbsp;&nbsp;不通过<input type="radio" name="stat" value=0  style="border:none"
		<%
		If Cint(Ds(6,0))=0 Then Response.Write "checked"
		%>
		></td>
      </tr>	
	 <tr>
<td height="30">&nbsp;</td>
<td class="category">
<input type="hidden" name="action" value="review">
<input type="hidden" name="id" value=<%=Ds(0,0)%>>
<input type="submit" value=" 确认添加 " class="button">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value=" 重新填写 " class="button">
</td>
</tr>
		</table>
		</form>
</td>
<td></td>
</tr>
<tr>
<td><img src="../skin/images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="../skin/images/r_3.gif" alt="" /></td>
</tr>
</table>
<% End If %>
</body>
</html>
<%
End Sub
%>
