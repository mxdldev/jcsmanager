<!-- #include file="Const.Asp" -->

<% '强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

'----------定义导航
Dim Navigation
Navigation = SystemName & ">> 报纸阅读" &">> 所有报纸"
'----------判断状态
Dim action
action = LCase(Request("actiont"))
Select Case action
Case "search"     '查询
Navigation = SystemName & ">> 报纸阅读" &">> 报纸查询"
Call MainSub()
Case Else 
Call MainSub()
End Select 
'-------------------------报纸列表主窗体部分开始
Sub MainSub()
%>

<html>
<head>
<title>报纸阅读</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>

<BODY>
<%  
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField
 Dim publishDate_Array,publishDate_G,path_Open,visitTimeKeyWord,parameter_now
 '--------------------------------------------
'搜索条件+++++++++++++++++++++++++++++++++++++
visitTimeKeyWord=request("visitTimeKeyWord")
 keyword = Jcs.Checkstr(Request("keyword"))
If  visitTimeKeyWord<>"" Then
parameter_now=parameter_now&"&visitTimeKeyWord="&visitTimeKeyWord
keystr = "deleted = 0  And  PaperName Like  '%"&keyword&"%' And publishDate = '"&visitTimeKeyWord&"'"
Else 
keystr = "deleted = 0  And  PaperName Like  '%"&keyword&"%' "
End If 

OrderBy = 1
OrderField = "PaperID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))

If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If
'+++++++++++++++++++++++++++++++++++++++++++++++++
 %>
<script>
  function Show(path)
{
window.open(path,'new','fullscreen=yes');
}
</script>
<script src="../skin/js/WdatePicker.js" language="javascript"></script>
<% '----检查阅读员权限(取消权限限制：所以用户都可以在线阅读)
  'Jcs_Manger.Manageckeck()

'取得当前页码
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">

  <input type="hidden" name="form" value="form2">
  <input type="hidden" name="actiont" value="search">
  <tr> 
    <td width="39%" height="30">&nbsp;</td>
	<td width="56%" align="right">
	  <div align="right">报纸查询：日期 
	    <input type="text" class="Wdate" id="d11"  name="visitTimeKeyWord"   onClick="WdatePicker()" onFocus="WdatePicker({onclearing:function(){if(!confirm('日期框的值为:'+this.value+', 确实要清空吗?'))return true;}})"/>
	    <input type="text" name="keyword" size="20" value="">
	    <input type="submit" value=" 查询 " class="button">
  &nbsp;&nbsp;</div></td>
    <td width="5%" align="right">&nbsp;</td>
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
	   <a href="?action=order&orderid=PaperID&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PaperID&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="33%" class="category">
	  报纸名称
	  <a href="?action=order&orderid=Tb_OfficeName&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=Tb_OfficeName&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
     <td width="14%" class="category">
	 出版日期
	   <a href="?action=order&orderid=PublishDate&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PublishDate&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
    <td class="category" width="24%">所属报社</td>
<td class="category" width="9%">在线阅读</td>
  </tr>
  <% 

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 2
SQL = "Select Count(PaperID) From [Tb_Paper] Where deleted = 0  And "& keystr
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>没有找到报纸</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "v_Paper"
Cmd("@Fields") = "PaperID,PaperName,PublishDate,Tb_OfficeName"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)
'----------------
 publishDate_Array=split(Ds(2,i),"-")
  if len(publishDate_Array(1))=1 then
  publishDate_Array(1)="0"&publishDate_Array(1)
  end if
    if len(publishDate_Array(2))=1 then
  publishDate_Array(2)="0"&publishDate_Array(2)
  end if
  publishDate_G=publishDate_Array(0)&publishDate_Array(1)&publishDate_Array(2)
  path_Open="/SLB/"&publishDate_G
'----------------
Response.Write("<tr onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''"" onDblClick=""Show('"&path_Open&"');"" ><td align=""center"" height=""25"">"&Ds(0,i)&"</td><td align=""center"">"&Ds(1,i)&"</td><td align=""center"">"&Ds(2,i)&"</td>")
Response.Write("<td align=""center"">"&Ds(3,i)&"</td>")
Response.Write("<td align=""center""><a href=""javascript:void(0);"" onClick=""Show('"&path_Open&"');"">点击查看</a></td></tr>")

 Next:Response.Flush

' Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">&nbsp;
	  双击每行可以在线阅读报纸</td>  
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

 %>