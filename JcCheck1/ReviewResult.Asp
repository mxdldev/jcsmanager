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
Dim action,Navigation,PaperName,PublishDate
action = LCase(Request("action"))
PaperName = Jcs.Checkstr(Request("PaperName"))
PublishDate=Jcs.Checkstr(request("PublishDate"))
Navigation = SystemName & ">> 报纸审读管理" &">>审核统计"

Select Case action
Case "del"      '删除
Call DelPaper()
Case Else 
Call MainSub()
End Select 
'-------------------------报社列表主窗体部分开始
Sub MainSub()
 Dim Ds,i,Count,Page,Pagesize,Cmd,MaxPage,keystr,OrderBy,OrderField,PaperDir,Check,stated,FieldName(1),FieldValue(1)
 %>

<html>
<head>
<title>审核统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script src="../skin/js/paper.js" language="javascript"></script>
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
	  报纸查询(报纸名称)：
	    <input  type="hidden" name="action" value="search"> 
	    <input type="text" name="PaperName" size="20" value="">
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
    <td width="10%" height="30" class="category">编号<a href="?action=order&orderid=PaperID&Orderby=0&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PaperID&Orderby=1&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
<td width="10%" class="category">报纸名称</td><td width="15%" class="category">出版日期<a href="?action=order&orderid=PublishDate&Orderby=0&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PublishDate&Orderby=1&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td><td width="10%" class="category">文章总数</td><td width="10%" class="category">未审核</td><td width="10%" class="category">初审通过</td><td width="10%" class="category">初审未过</td><td width="10%" class="category">预审通过</td><td width="10%" class="category">预审未过</td><td width="10%" class="category">图片文章</td><td width="10%" class="category">图片未审</td><td width="10%" class="category">图片通过</td><td width="10%" class="category">图片未过</td>
   <td width="10%" class="category">状态</td>
  </tr>
  <% 
If PublishDate="" Then
  FieldName(0)="PaperName"
  FieldValue(0)=PaperName
keystr = "Deleted = 0 And UserID="&Jcs.UserID&" And PaperName Like  '%"&PaperName&"%'"
Else
  FieldName(0)="PaperName"
  FieldValue(0)=PaperName
  FieldName(1)="PublishDate"
  FieldValue(1)=PublishDate
keystr = "Deleted = 0 And UserID="&Jcs.UserID&" And PublishDate='"&PublishDate&"' And PaperName Like  '%"&PaperName&"%'"
End If
'-----
OrderBy = 1
OrderField = "PaperId"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))
'Response.Write(Request.QueryString("orderby"))
'Response.Write(Request.QueryString("orderid"))
If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 3
SQL = "Select Count(PaperID) From [Tb_Paper] Where "& keystr

'response.Write sql
'response.End()
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>暂时没有相关的报纸</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "Tb_Paper"
Cmd("@Fields") = "PaperId,PaperName,PaperType,PublishDate,Folder,Tb_Check"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)
Response.Write "<tr align=""center""><td>"&Ds(0,i)&"</td><td>"&Ds(1,i)&"</td><td>"&Ds(3,i)&"</td><td>"&GetAll(Ds(0,i),100)&"</td><td>"&GetAll(Ds(0,i),0)&"</td><td>"&GetAll(Ds(0,i),1)&"</td><td>"&GetAll(Ds(0,i),2)&"</td><td>"&GetAll(Ds(0,i),3)&"</td><td>"&GetAll(Ds(0,i),4)&"</td><td>"&GetAll(Ds(0,i),5)&"</td><td>"&GetAll(Ds(0,i),6)&"</td><td>"&GetAll(Ds(0,i),7)&"</td><td>"&GetAll(Ds(0,i),8)&"</td><td style=""color:#f00"">"&vbcrlf
Check=Cint(Ds(5,i))
If Cint(GetAll(Ds(0,i),100))=Cint(GetAll(Ds(0,i),3)) And Cint(GetAll(Ds(0,i),5))=Cint(GetAll(Ds(0,i),7)) And Check<>3 Then 
Check=3
UpdateCheck(Ds(0,i))
End If
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
  Response.Write "</td>"
 Next:Response.Flush
' Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="15" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">
</td>
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
    <td width="20%" align="right"> <input type="hidden" name="page" value="<%=Request("Page")%>">
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

Function GetAll(PaperId,vType)
  Dim SQL
  Select Case vType
  Case 100
  SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID
  Case 5
  SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID&" And (Images<>'' or Images<>Null)"
   Case 6
  SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID&" And (Images<>'' or Images<>Null) And ImageChecks=0 "
  Case 7
  SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID&" And ImageChecks=1"
  Case 8
     SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID&" And ImageChecks=2"
  Case Else
  SQL = "Select Count(ArticleID) From [Tb_Article] Where Deleted = 0 And PaperId="&PaperId&" And UserID="&Jcs.UserID&" And ArticleChecks="&vType
  End Select
GetAll = Jcs.Execute(SQL)(0)
End Function

Sub UpdateCheck(PaperId)
  Dim Sql
  Sql="Update Tb_Paper Set Tb_Check=3 Where PaperId="&PaperId
  Jcs.Execute(Sql)
End Sub
Sub DelPaper()
Dim ID,i,J,page
page = Jcs.Checkstr(Request("page"))
J = 0
ID =  Jcs.Checkstr(Request("ID"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Tb_Paper] Set Deleted =  1 Where PaperID = " & ID(i)
Jcs.Execute(SQL)
J = J + 1
Next 
Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 
 %>