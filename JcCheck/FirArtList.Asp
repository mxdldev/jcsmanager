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
Dim action,Navigation,Title,PublishDate,PaperId
action = LCase(Request("action"))
PaperId=Jcs.Checkstr(Request("PaperId"))
Title=Jcs.Checkstr(Request("Title"))
PublishDate=Jcs.Checkstr(request("PublishDate"))
If action="search" or Title<>"" or PublishDate<>"" Then
Navigation = SystemName & ">> 报纸审读管理" &">>报纸预审>>文章查询"
Else
Navigation = SystemName & ">> 报纸审读管理" &">>报纸预审>>文章列表"
End If

Select Case action
Case "del"      '删除
Call Jcs_Check.DelArticle()
Case Else 
Call MainSub()
End Select 
'-------------------------报社列表主窗体部分开始
Sub MainSub()
 Dim Ds,i,Count,Page,Pagesize,Cmd,MaxPage,keystr,OrderBy,OrderField,PaperDir,Check,stated,FieldName(2),FieldValue(2)
 %>

<html>
<head>
<title>文章浏览</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script src="../skin/js/article.js" language="javascript"></script>
<script src="../skin/js/calenda1.js" language="javascript"></script>
</HEAD>

<BODY>

<% 
'----检查管理员权限
 ' Jcs_Manger.Manageckeck()

'取得当前页码
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form action=""  name="SearchForm" >
  <tr> 
    <td width="37%" height="30">&nbsp;</td>
	<td width="63%" align="left">
	  文章查询(文章标题)：
	    <input  type="hidden" name="action" value="search"> 
		<input type="hidden" name="PaperId" value="<%=PaperId%>">
	    <input type="text" name="Title" size="20" value="">
	  <input type="submit" value=" 查询 " class="button" onClick="return CheckInput()">&nbsp;出版日期：<span style="color:#FF0000;">
	  <%= PublishDate %></span>
      <input name="PublishDate" type="hidden" class="InputTxt" size="11"  readonly ><A onclick=event.cancelBubble=true; href="javascript:showCalendar('imageCalendar1',false,'PublishDate',null,'',0);"> <IMG id="imageCalendar1" height=21 src="../skin/images/button.gif" width=34 align="absMiddle" border="0"></A></td>
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
<form name="myform" onSubmit="return ConfirmDel();">

<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1">
  <tr align="center">
    <td width="5%" class="category">选中</td>
    <td width="10%" class="category">编号<a href="?action=order&orderid=ArticleID&Orderby=0&PaperId=<%=PaperId%>&PublishDate=<%=PublishDate%>&Title=<%=Title%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=ArticleID&Orderby=1&PaperId=<%=PaperId%>&PublishDate=<%=PublishDate%>&Title=<%=Title%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="15%"class="category">报纸名称</td>
    <td width="15%" class="category">发布时间<a href="?action=order&orderid=PublishDate&Orderby=0&PaperId=<%=PaperId%>&PublishDate=<%=PublishDate%>&Title=<%=Title%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PublishDate&Orderby=1&PaperId=<%=PaperId%>&PublishDate=<%=PublishDate%>&Title=<%=Title%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="20%" class="category">文章标题</td>
	<td width="10%" class="category">版名</td>
	<td width="5%" class="category">版次</td>
    <td width="10%" class="category">状态</td>
    <td  width="10%" class="category">操作</td>
  </tr>
  <% 
If PublishDate="" Then
  FieldName(0)="PaperId"
  FieldValue(0)=PaperId
  FieldName(1)="Title"
  FieldValue(1)=Title
keystr = "Deleted = 0 And ArticleChecks between 1 And 2 And PaperId="&PaperId&" And Title Like  '%"&Title&"%'"
Else
  FieldName(0)="PaperId"
  FieldValue(0)=PaperId
  FieldName(1)="Title"
  FieldValue(1)=Title
  FieldName(2)="PublishDate"
  FieldValue(2)=PublishDate
keystr = "Deleted = 0 And ArticleChecks between 1 And 2 And PaperId="&PaperId&" And PublishDate='"&PublishDate&"' And Title Like  '%"&Title&"%'"
End If
'-----
OrderBy = 1
OrderField = "ArticleId"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))
'Response.Write(Request.QueryString("orderby"))
'Response.Write(Request.QueryString("orderid"))
If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 8
SQL = "Select Count(ArticleID) From [Tb_Article] Where "& keystr

'response.Write sql
'response.End()
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>暂时没有要预审的文章</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "Tb_Article"
Cmd("@Fields") = "ArticleId,PaperId,VerName,PublishDate,VerOrder,Title,Author,ArticleChecks"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)
Response.Write "<tr align=""center""><td><input name=""ID"" type=""checkbox"" id=""ID"" value='"&Ds(0,i)&"' style=""border:0""></td><td>"&Ds(0,i)&"</td><td>"&Jcs_Check.GetPaperName(Ds(1,i))&"</td><td>"&Ds(3,i)&"</td><td>"&Ds(5,i)&"</td><td>"&Ds(2,i)&"</td><td>"&Ds(4,i)&"</td><td style=""color:#f00"">"&vbcrlf
Check=Cint(Ds(7,i))
Select Case Check
Case 0
 stated="未审核"
Case 1
 stated="初审通过"
Case 2
 stated="初审未过"
Case 3
 stated="终审通过"
End Select
 response.write stated
 Check=0
  Response.Write "</td><td><a href=""SecReview.asp?action=show&id="&Ds(0,i)&""">预审</a>|<a href=""?action=del&id="&Ds(0,i)&"&PaperId="&PaperId&"&Page="&Page&""" onclick=""return ConfirmDel()"">删除</a></td></tr>"
'Response.Write "<tr align=""center""><td>"&Ds(0,i)&"</td><td>"&Ds(1,i)&"</td><td>"&Ds(2,i)&"</td><td>"&Ds(3,i)&"</td><td><a href="&PaperDir&">查看</a>|删除</td></tr>"
 Next:Response.Flush
 Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">
全选：<input name="chkAll" type="checkbox" id="chkAll" value="" onClick="CheckAll(this.form)" style="border:0"> <input name="Action" type="hidden" id="Action" value="Del"><input name="submit" type="submit" class="button"  value="批量删除选定的文章"></td>
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
    <td width="20%" align="right"><input type="hidden" name="page" value="<%=Request("Page")%>">
		<input type="hidden" name="PaperId" value="<%=PaperId%>"><input type="hidden" name="PublishDate" value="<%=PublishDate%>">
		<input type="hidden" name="Title" value="<%=Title%>">
		</td>
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
 %>