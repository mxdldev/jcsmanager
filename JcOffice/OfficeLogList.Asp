<!-- #include file="../JcManger/Const.Asp" -->
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
Navigation = SystemName & ">> 日志查看" &">> 我的日志"
'----------判断状态
Dim action
action = LCase(Request("actiont"))
Select Case action
Case "search"     '查询
Navigation = SystemName & ">> 日志查看" &">> 我的日志"
Call MainSub()
Case Else 
Call MainSub()
End Select 
'-------------------------日志列表主窗体部分开始
Sub MainSub()
%>
<html>
<head>
<title>审读员日志</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
</HEAD>
<BODY>
<% '----检查管理员权限
'Jcs.CheckManger()'取得当前页码
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <input type="hidden" name="form" value="form2">
  <input type="hidden" name="actiont" value="search">
  <tr> 
    <td width="5%" height="30">&nbsp;</td>
	<td width="95%" align="right">
	  日志查询：
	    
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
    <td width="14%" height="30" class="category">
	 Ip</td>
    <td width="14%" class="category">
	  操作动作 </td><td width="37%" class="category"> 日志内容</td>
    <td class="category" width="17%">操作时间
	<a href="?action=order&orderid=LogDate&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=LogDate&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>

    <td class="category" width="8%">操作用户</td>
	
	<td class="category" width="10%">查看</td>
  </tr>
  <% 
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField,trcolor
 
 keyword = Jcs.Checkstr(Request("keyword"))
keystr = "DeleteFlag = 0 And UserClassID = 3  And UserID="&Jcs.UserID&" And (LogType Like '%"&keyword&"%' Or LogContent Like '%"&keyword&"%'  Or  Ip Like  '%"&keyword&"%'  Or  TrueName Like  '%"&keyword&"%' Or OfficeName Like  '%"&keyword&"%' )"
'-----
OrderBy = 1
OrderField = "LogID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))
If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If
Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 12
SQL = "Select Count(LogID) From [v_Log] Where DeleteFlag = 0  And "& keystr
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>没有找到您要查询的日志！</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "v_Log"
Cmd("@Fields") = "LogID,IP,LogType,LogContent,LogDate,TrueName,UserClassID"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
trcolor ="#fff"
For i = 0 To UBound(Ds,2)
Response.Write("<tr onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''"" onDblClick=""javascript:var win=window.open('LogView.Asp?id="&Ds(0,i)&"','日志详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()"" ><td align=""center"" height=""25"">"&Ds(1,i)&"</td><td align=""center"">"&Ds(2,i)&"</td><td align=""center""><a href=""LogView.Asp?id="& Ds(0,i)& """ >"&Jcs.cutString(Ds(3,i),42)&"</a></td>")
Response.Write("<td align=""center"">"&Ds(4,i)&"</td>")
Response.Write("<td align=""center"">"&Ds(5,i)&"</td><td align=""center""><a href=""LogView.Asp?id="& Ds(0,i)& """>点击查看</a></td></tr>")
Next:Response.Flush
' Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr >
	<td width="49%" align="left" style="color:#FF0000;" >&nbsp;
	  双击每行可查看日志详细资料
	  </td>  
	<td width="46%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
	 
    <td width="5%" align="right">&nbsp;</td>
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