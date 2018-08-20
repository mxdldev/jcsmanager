<!-- #include file="../Inc/Const.Asp" -->
<!-- #include file="Function.asp" -->
<% 
'判断登陆
Jcs.CheckManger()
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报刊管理" &">> 所有报刊"
'----------判断状态
Dim action
action = LCase(Request("actiont"))
Select Case action
Case "del"
Call DelOffice()
Case "order"      '排序
Call OrderOffice()
Case "search"     '查询
Navigation = SystemName & ">> 报刊管理" &">> 报刊查询"
Call MainSub()
Case Else 
Call MainSub()
End Select 
'-------------------------报社列表主窗体部分开始
Sub MainSub()
%>

<html>
<head>
<title>报社查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<script>
function CheckAll(form)  {
  for   (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<% '----检查管理员权限
Call  Manageckeck()

'取得当前页码
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <input type="hidden" name="form" value="form2">
  <input type="hidden" name="actiont" value="search">
  <tr> 
    <td width="5%" height="30">&nbsp;</td>
	<td width="95%" align="right">
	  报刊查询：
	    
	    <input type="text" name="keyword" size="20" value="">
	  <input type="submit" value=" 查询 " class="button">&nbsp;	</td>
  </tr>
</form>  
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../skin/images/r_1.gif" alt="" /></td>
<td width="100%" background="../skin/images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;<%= Navigation %></td>
	  <td align="right">&nbsp;</td>
    </tr>
  </table>
</td>
<td><img src="../skin/images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td></td>
<td>
<form name="form1" action="Office_Del.Asp">
  <input type="hidden" name="form" value="form1">

<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">


  <tr align="center">
    <td width="11%" height="30" class="category">
	 编号
	   <a href="?action=order&orderid=Tb_ClassID&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=Tb_ClassID&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="33%" class="category">
	  报刊名称
	  <a href="?action=order&orderid=Tb_ClassName&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=Tb_ClassName&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
  
	<td width="25%" class="category">
	 报刊名称
	   <a href="?action=order&orderid=Tb_OfficeID&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=Tb_OfficeID&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
	

    <td class="category" width="9%">修改</td>

	
    <td class="category" width="13%">删除</td>
	
	
	<td class="category" width="9%">查看</td>
  </tr>
  <% 
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField
 
 keyword = Jcs.Checkstr(Request("keyword"))
keystr = " Tb_ClassName Like  '%"&keyword&"%'"
'-----
OrderBy = 1
OrderField = "Tb_ClassID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))
Response.Write(Request.QueryString("orderby"))
Response.Write(Request.QueryString("orderid"))
If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request.QueryString("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 2
SQL = "Select Count(Tb_ClassID) From [Tb_Class] Where DeleteFlag = 0  And "& keystr
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>没有找到报刊</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "Tb_CLass"
Cmd("@Fields") = "Tb_ClassID,Tb_ClassName,Tb_OfficeID"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)

Response.Write("<tr onDblClick=""javascript:var win=window.open('Office_View.Asp?id="&Ds(0,i)&"','报社详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()""><td align=""center"" height=""25"">"&Ds(0,i)&"</td><td align=""center"">"&Ds(1,i)&"</td><td align=""center"">"&FindOffice(Ds(2,i))&"</td>")
Response.Write("<td align=""center""><a href=""TbClass_Edit.Asp?action=edit&id="&Ds(0,i)&"""><img src=""../skin/images/res.gif"" border=""0"" hspace=""2"" align=""absmiddle"">修改</a></td>")
Response.Write("<td align=""center""><input type=""checkbox"" name=""ID"" value="""" style=""border:0""></td><td align=""center""><a href=""Press_View.Asp?id="& Ds(0,i)& """>点击查看</a></td></tr>")

 Next:Response.Flush

' Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">&nbsp;
	  <%if request("form")<>"" then%>
	双击选择报刊
	  <%else%>
	  双击每行可查看报刊详细资料
	  <%end if%></td>  
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
    <td width="20%" align="right"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0">
      全选
        <input name="submit" type="submit" class="button" onClick="return confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除所选择的会员吗？')" value="删 除"></td>
	</tr>
 
  </table></td></tr>
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
 <% end if %>
</table>
</body>
</html>
<% 
End Sub 

'----------删除报刊
Sub DelOffice()
Dim ID,i,J,page,pressname
page = Jcs.Checkstr(Request("page"))
J = 0
ID =  Jcs.Checkstr(Request("ID"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Tb_Class] Set DeleteFlag =  1 Where Tb_ClassID = " & ID(i)
Response.Write sql
Jcs.Execute(SQL)
J = J + 1
'---写操作日志
Jcs_Manger.GetOffice ID(i)
pressname = Jcs_Manger.Tb_ClassName
Jcs.WriteLog pressname,"删除报社"
Jcs_Manger.ClearOffice
Next 

'Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 

'---------------------检查管理员权限
Sub Manageckeck()
If Jcs.UserClassID <> 3 Then 
Response.Write("<br><center><img src=""../skin/images/note.gif"" align=""absmiddle"">&nbsp;<font color=""#FF0000"">你不具备此权限，请与管理员联系！</font></center>")
Response.end
End If
End Sub 

 %>