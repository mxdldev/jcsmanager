<!-- #include file="Const.Asp" -->
<!-- #include file="../Inc/MD5.Asp" -->
<script language="javascript" type="text/javascript" src="../skin/Js/Jcs_Function.js"></script>
<% '强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'判断登陆
'----------定义导航
Dim Navigation,Navigation2
Navigation = SystemName & ">> 报社用户管理" &">> 所有报社用户"
Navigation2 = SystemName & ">> 报社用户管理" &">> 添加报社用户"
'----------判断状态
Dim action
action = LCase(Request("actiont"))
Select Case action
Case "add"
Call AddUser()
Case "lock"     '锁定
Call LockUserInfo()
Case "search"     '查询
Navigation = SystemName & ">> 报社用户管理" &">> 报社用户查询"
If Request("actiondel") = "sub" Then 
Call DelUser()	
Else 
Call MainSub()
End If 
Case "sub"      '删除
Call DelUser()
Case Else 
If Request("actiondel") = "sub" Then 
Call DelUser()	
Else 
Call MainSub()
End If 
End Select 
'-------------------------报社用户列表主窗体部分开始
Sub MainSub()
%>

<html>
<head>
<title>报社用户管理</title>
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
  function check()
{
var patrn=/^[a-zA-Z][a-zA-Z0-9_]{4,15}$/;
var parpwdmac=/^(\w){6,20}$/;
if (document.formadd.officename.value=="")
{
alert("请选择报社！");
document.formadd.officename.focus();
return false;
}
if (document.formadd.username.value=="")
{
alert("用户名称没有填写！");
document.formadd.username.focus();
return false;
}
if (!patrn.exec(document.formadd.username.value))
{
alert("用户名格式不正确！");
document.formadd.username.focus();
return false;
}
if (document.formadd.password.value=="")
{
alert("密码没有填写！");
document.formadd.password.focus();
return false;
}
if (!parpwdmac.exec(document.formadd.password.value))
{
alert("密码格式不正确！");
document.formadd.password.focus();
return false;
}
if (document.formadd.repassword.value!=document.formadd.password.value)
{
alert("两次密码不一样，请重新输入！");
document.formadd.repassword.focus();
return false;
}
if (document.formadd.truename.value=="")
{
alert("真实姓名没有填写！");
document.formadd.truename.focus();
return false;
}
}
//
</script>
<% '----检查管理员权限
 ' Jcs_Manger.Manageckeck()
%>
<form name="formadd" ACTION="" METHOD="post" >
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../skin/images/r_1.gif" alt="" /></td>
<td width="100%" background="../skin/images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;<%= Navigation2%></td>
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
        <td width="25%" height="30" align="right">用户名：</td>
        <td width="75%" class="category"><input type="text" name="username" style="width:200px" id="Username" onKeyUp="javascript:document.getElementById('Submit').disabled = true;"  Onblur="ajaxLoadPage('ajaxdata.asp?action=user&ID=',this.value,'get','IsOkUserName','1.1.1');" >
          &nbsp;&nbsp;<span id="IsOkUserName">(字母开头，允许5-16字节，允许字母数字下划线)</span></td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">密码：</td>
        <td width="75%" class="category"><input  type="password" name="password" style="width:200px">
          &nbsp;&nbsp;(6-20个字母、数字、下划线)</td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">确认密码：</td>
        <td width="75%" class="category"><input  type="password" name="repassword" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	
	   <tr>
        <td width="25%" height="30" align="right">真实姓名：</td>
        <td width="75%" class="category"><input type="text" name="truename" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	  
      <tr>
	    <td height="30">&nbsp;</td>
        <td class="category">
		 <input type="hidden" name="actiont" value="add">
		  <input type="submit" value=" 确认添加 " name="submit" id="submit" onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
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
<script language="javascript" defer="defer">
document.getElementById('Submit').disabled = true;
</script>
</form>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <input type="hidden" name="form" value="form2">
  <input type="hidden" name="actiont" value="search">
  <tr> 
    <td width="5%" height="30">&nbsp;</td>
	<td width="95%" align="right">
	  报社用户查询：
	    
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
    <td width="10%" height="30" class="category">
	 编号
	   <a href="?action=order&orderid=UserID&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=UserID&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="8%" class="category">
	  用户姓名	  </td>
  
	<td width="16%" class="category">
	 最后登陆时间   <a href="?action=order&orderid=LastLoginTime&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=LastLoginTime&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
	 <td width="23%" class="category">
	 所属报社</td>
	 <td width="8%" class="category">
	 用户状态</td>
	

    <td class="category" width="9%">修改</td>

	
    <td class="category" width="5%">删除</td>
	
	
	<td class="category" width="8%">查看详细</td>
  </tr>
  <% 
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField,officeid
officeid = GetOfficeid(Jcs.UserID)
 
keyword = Jcs.Checkstr(Request("keyword"))
keystr = "UserClassID = 3 And (Tb_OfficeId = "&officeid&") And (TrueName Like  '%"&keyword&"%' Or UserName Like  '%"&keyword&"%' )"
'-----
OrderBy = 1
OrderField = "UserID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))

If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 4
SQL = "Select Count(UserID) From [v_JcUser] Where 1 = 1  And "& keystr
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>没有找到报社用户</b></td></tr>"
Else
MaxPage = Count\Pagesize
If Count Mod Pagesize <> 0 Then MaxPage = MaxPage + 1
If Int(Page) > MaxPage Then Page = MaxPage
Set Cmd = Server.CreateObject("ADODB.Command")
Set Cmd.ActiveConnection=Conn
Cmd.CommandText="PageList"
Cmd.CommandType=4
Cmd.Prepared = True
Cmd("@Tablename") = "v_JcUser"
Cmd("@Fields") = "UserID,TrueName,UserName,LastLoginTime,UofficeName,UserLevel,cndeleteflag"
Cmd("@OrderBy") = OrderField
Cmd("@Pagesize")= Pagesize
Cmd("@PageIndex") = Page
Cmd("@OrderType") = OrderBy
Cmd("@Where") = keystr
Set Rs = Cmd.Execute
If Not (Rs.Bof And Rs.Eof) Then Ds = Rs.GetRows(-1)
Rs.Close:Set Rs = Nothing
For i = 0 To UBound(Ds,2)

Response.Write("<tr onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''"" onDblClick=""javascript:var win=window.open('User_View.Asp?id="&Ds(0,i)&"','报社用户详细信息','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()"" ><td align=""center"" height=""25"">"&Ds(0,i)&"</td><td align=""center"">"&Ds(1,i)&"</td><td align=""center"">"&Ds(3,i)&"</td><td align=""center"">"&Ds(4,i)&"</td><td align=""center"">"&LockStat(Ds(0,i))&"</td>")
Response.Write("<td align=""center""><a href=""User_Edit.Asp?url=OfficeUsers_List.Asp&id="& Ds(0,i)&"&page="&Page& """><img src=""../skin/images/res.gif"" border=""0"" hspace=""2"" align=""absmiddle"">修改</a></td>")
Response.Write("<td align=""center""><input type=""checkbox"" name=""ID"" value=" &Ds(0,i)& " style=""border:0""></td><td align=""center""><a href=""User_View.Asp?url=OfficeUsers_List.Asp&id="& Ds(0,i)&"&page="&Page& """>点击查看</a></td></tr>")

 Next:Response.Flush

' Closedatabase()   '关闭数据库链接
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">&nbsp;
	  双击每行可查看用户详细资料; 状态：点击"正常"则该用户被锁定。
	  </td>  
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
	 
    <td width="20%" align="right"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0"> <input type="hidden" name="page" value="<%=Request("Page")%>">
      全选 <input type="hidden" name="actiondel" value="sub">
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

'----------删除报社用户
Sub DelUser()
Dim ID,i,J,page,Username
page = Jcs.Checkstr(Request("page"))
J = 0
ID =  Jcs.Checkstr(Request("ID"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Jc_User] Set DeleteFlag =  1 Where NID = " & ID(i)
Jcs.Execute(SQL)
J = J + 1
'---写操作日志
Jcs_Manger.GetUser ID(i)
Username = Jcs_Manger.UserTitle
Jcs.WriteLog Username,"删除报社用户"
Jcs_Manger.ClearUser
Next 

'Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 
'-------------------添加报社用户
Sub AddUser()
Dim Uname,Pwd,Tname,Tbofficeid,ND
Uname = Jcs.Checkstr(Request.Form("username"))
Pwd = Jcs.Checkstr(Request.Form("password"))
Tname = Jcs.Checkstr(Request.Form("truename"))
Tbofficeid = GetOfficeid(Jcs.UserID)
If Uname = ""   Then
	Founderr = True
	Msg = "用户名不能为空！"
	Jcs.BackInfo "",0
	Exit Sub
End If
'--------------------判断用户是否存在
If Not Founderr Then
	SQL = "Select UserName From [Jc_User] Where UserName = '"&Uname&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Founderr = True
	Rs.Close:Set Rs = Nothing     '释放Rs
		Msg = "用户名已经存在！"
		Jcs.BackInfo "",0
	End If 
End If 
If Not Founderr Then
ND = Jcs.CreatePass
Pwd= ND & MD5(ND & Pwd)
response.write(Uname)
response.write(Tname)
response.write(Tbofficeid)
response.write(Pwd)
SQL = "Insert Into  [Jc_User] (UserName,TrueName,Tb_OfficeId,UserPassWord,UserClassID) Values ('"&Uname&"','"&Tname&"',"&Tbofficeid&",'"&Pwd&"',3)"
Jcs.Execute(SQL)
'---写操作日志
Jcs.WriteLog Tname,"添加报社用户"
	Closedatabase()   '关闭数据库链接
Msg = "报社用户“" &Tname& "”添加成功！"
		Jcs.BackInfo "OfficeUsers_List.Asp",2
End If 
End Sub 

'------------------判读锁定状态
Function LockStat(ByVal id)
Dim LockStr
 Jcs_Manger.GetUserInfo(id)
LockStr = Jcs_Manger.UDeleteFlag

Select Case LockStr
Case 1
LockStat = "<a href=?actiont=lock >已删除</a>"
Case 2
LockStat = "<a href=?actiont=lock&ID="& id &"&page="&Request("Page")&" title=点击恢复使用>已锁定</a>"
Case Else 
LockStat = "<a href=?actiont=lock&ID="& id &"&page="&Request("Page")&" title=点击锁定>正常</a>"
End Select 
End Function
'-------------------锁定用户或开通用户
Sub LockUserInfo()
Dim id,delflag,mothed,truename,page
page = Jcs.Checkstr(Request("page"))
id =  Jcs.Checkstr(Request("ID"))
If IsNull(id)  Then 
 Msg = "参数错误"
 Founderr = True
Jcs.BackInfo "",0
End If 

Jcs_Manger.GetUserInfo id
delflag = Jcs_Manger.UDeleteFlag
Select Case delflag
Case 0
SQL = "Update [Jc_User] Set DeleteFlag =  2 Where UserID = " & id
mothed = "锁定用户"
Case 2
SQL = "Update [Jc_User] Set DeleteFlag =  0 Where UserID = " & id
mothed = "解除用户锁定"
Case Else
Founderr = True
End Select 
If Not Founderr Then 
Jcs.Execute(SQL)
'---写操作日志
truename = Jcs_Manger.TrueName
Jcs.WriteLog truename,mothed
Jcs_Manger.ClearUserInfo
 Msg = mothed &"”" &truename&"“成功！" 
Else 
 Msg = "该用户已经被删除"
End If 
  Jcs.BackInfo "?page="&page,2
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