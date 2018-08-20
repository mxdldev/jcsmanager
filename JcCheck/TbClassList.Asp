<!-- #include file="Const.Asp" -->
<% '强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'判断登陆
Jcs.CheckManger()
'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报刊管理" &">> 所有报刊"
'----------判断状态
Dim action
action = LCase(Request("action"))
Select Case action
Case "search"     '查询
Navigation = SystemName & ">> 报刊管理" &">> 报刊查询"
Case "del"
Call DelTbClass()
Case "review"
Call Jcs_Check.ClassReview()
Case Else
Call MainSub()
End Select 
'-------------------------公告列表主窗体部分开始
Sub MainSub()
%>

<html>
<head>
<title>报刊审核</title>
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
'Jcs_Manger.Manageckeck()

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
<form name="form3" ACTION="" METHOD="post" >
 
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td width="11%" height="30" class="category">
	 编号
	   <a href="?action=order&orderid=Tb_ClassId&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=Tb_ClassId&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="33%" class="category">
	  报刊名称</td>
  
	<td width="25%" class="category">
	 添加时间
	   <a href="?action=order&orderid=RegDate&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=RegDate&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
    <!--td class="category" width="13%">删除</td-->
	 <td class="category" width="9%">审核状态</td>
	 <td class="category" width="9%">操作</td>
	 <!--td class="category">操作</td-->
  </tr>
  <% 
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField,check,stated
 
 keyword = Jcs.Checkstr(Request("keyword"))
keystr = "DeleteFlag = 0  And  Tb_ClassName Like  '%"&keyword&"%'"
'-----
OrderBy = 1
OrderField = "Tb_ClassID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))

If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 2
SQL = "Select Count(Tb_ClassId) From [Tb_Class] Where "& keystr
'Response.write sql
'Response.end()
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
Cmd("@Tablename") = "Tb_Class"
Cmd("@Fields") = "Tb_ClassId,Tb_ClassName,Tb_Date,Tb_ClassState,Regdate"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)

Response.Write"<tr onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''"" onDblClick=""javascript:var win=window.open('TbClassReview.asp?id="&Ds(0,i)&"&action=show','报刊详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()"" ><td align=""center"" height=""25"">"&Ds(0,i)&"</td><td align=""center"">"&Ds(1,i)&"</td><td align=""center"">"&Ds(4,i)&"</td>"
'Response.Write("<td align=""center""><a href=""Notice_Edit.Asp?id="& Ds(0,i)& """><img src=""../skin/images/res.gif"" border=""0"" hspace=""2"" align=""absmiddle"">修改</a></td>")
Response.Write "<td align=""center"" style=""color:#f00"">"
If Cint(Ds(3,i))=0 Then
Response.Write"未通过"
Else
Response.Write "通过"
End IF
Response.Write "</td>"
'Response.Write "<td align=""center"">"
'If Cint(Ds(3,i))=0 Then
'Response.Write "<a href=""?action=review&stat=1&id="&Ds(0,i)&""">通过审核</a>"
'else
'Response.Write "<a href=""?action=review&stat=0&id="&Ds(0,i)&""">取消审核</a>"
'End If
'Response.Write "</td>"
Response.Write"<td align=""center""><a href=""TbClassReview.asp?id="& Ds(0,i)& "&action=show"">审核</a></td>"
Response.Write "</tr>"

 Next:Response.Flush

Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">&nbsp;
	  双击每行可查看报刊详细
	  </td>  
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
	 
    <td width="20%" align="right"></td>
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

'----------删除公告
Sub DelTbClass()
Dim ID,i,J,page,Noticename
page = Jcs.Checkstr(Request("page"))
J = 0
ID =  Jcs.Checkstr(Request("ID"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Tb_Class] Set DeleteFlag =  1 Where Tb_ClassID = " & ID(i)
Jcs.Execute(SQL)
J = J + 1
'---写操作日志

Jcs.WriteLog Jcs_Check.GetClassName(ID),"删除报刊"
Jcs_Manger.ClearNotice
Next 

'Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 


 %>