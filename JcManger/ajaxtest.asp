<!-- #include file="Const.Asp" -->
<!-- #include file="../Inc/MD5.Asp" -->
<% 'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'�жϵ�½
Jcs.CheckManger()
'----------���嵼��
Dim Navigation,Navigation2
Navigation = SystemName & ">> �����û�����" &">> ���б����û�"
Navigation2 = SystemName & ">> �����û�����" &">> ���ӱ����û�"
'----------�ж�״̬
Dim action
action = LCase(Request("actiont"))
Select Case action
Case "add"
Call AddUser()
Case "lock"     '����
Call LockUserInfo()
Case "search"     '��ѯ
Navigation = SystemName & ">> �����û�����" &">> �����û���ѯ"
If Request("actiondel") = "sub" Then 
Call DelUser()	
Else 
Call MainSub()
End If 
Case "sub"      'ɾ��
Call DelUser()
Case Else 
If Request("actiondel") = "sub" Then 
Call DelUser()	
Else 
Call MainSub()
End If 
End Select 
'-------------------------�����û��б������岿�ֿ�ʼ
Sub MainSub()
%>

<html>
<head>
<title>�����û�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
.black_overlay{
display:none;
position:absolute;
top:0%;
left:0%;
width:100%;
height:100%;
background-color:black;
z-index:1000;
opacity:.80;
filter:alpha(opacity=80);
}
#Layer1 {
font-family:verdana,tahoma,arial;
display:none;
padding:1px;
text-align:left;
position:absolute;
top:5%;
left:20%;
border:#5d7f9f 1px solid;
width:500px;
background-color:white;
z-index:1001;
}
#DivUserAM{
font-family:verdana,tahoma,arial;
display:none;
padding:1px;
text-align:left;
position:absolute;
	left:167px;
	top:34px;
	width:607px;
	height:189px;
	z-index:1002;
border:#5d7f9f 1px solid;
background-color:white;
z-index:1001;
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
alert("��ѡ���磡");
document.formadd.officename.focus();
return false;
}
if (document.formadd.username.value=="")
{
alert("�û�����û����д��");
document.formadd.username.focus();
return false;
}
if (!patrn.exec(document.formadd.username.value))
{
alert("�û�����ʽ����ȷ��");
document.formadd.username.focus();
return false;
}
if (document.formadd.password.value=="")
{
alert("����û����д��");
document.formadd.password.focus();
return false;
}
if (!parpwdmac.exec(document.formadd.password.value))
{
alert("�����ʽ����ȷ��");
document.formadd.password.focus();
return false;
}
if (document.formadd.truename.value=="")
{
alert("��ʵ����û����д��");
document.formadd.truename.focus();
return false;
}
}
//�õ�DOM��ID
function $(id)
{
	return document.getElementById(id);
}
//
function UserOperate()
{
	var cont_b=$("fade");
    cont_b.style.display='none';
	var obj=$("DivUserAM").style;
	obj.display="block"; 
}
//
function setviewdata(divid1,divid2)
{
var div_b=$(divid1);
var div_c=$(divid2);
 div_b.style.display='none';
 div_c.style.display='block';
}
</script>
<% '----������ԱȨ��
  Jcs_Manger.Manageckeck()
%>

<div id="fade" class="black_overlay"></div> 
<div id="DivUserAM">
<form name="formadd" ACTION="" METHOD="post" >
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#C4D8ED">
<tr>
<td><img src="../skin/images/r_1.gif" alt="" /></td>
<td width="100%" background="../skin/images/r_0.gif">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td>&nbsp;<%= Navigation2%></td>
	  <td align="right"><a href="javascript:void(0)" style="color:#000000" onClick="$('DivUserAM').style.display='none';$('fade').style.display='none';"><strong>�ر�</strong></a></td>
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
        <td width="13%" height="30" align="right">�������磺</td>
        <td width="87%" class="category"><%= Jcs.OfficeSelectlist(0)%></td>
      </tr>	
	  <tr>
        <td width="13%" height="30" align="right">�û�����</td>
        <td width="87%" class="category"><input type="text" name="username" style="width:200px">
          &nbsp;&nbsp;(��ĸ��ͷ������5-16�ֽڣ�������ĸ�����»���)</td>
      </tr>	
	   <tr>
        <td width="13%" height="30" align="right">���룺</td>
        <td width="87%" class="category"><input  type="password" name="password" style="width:200px">
          &nbsp;&nbsp;(6-20����ĸ�����֡��»���)</td>
      </tr>	
	   <tr>
        <td width="13%" height="30" align="right">��ʵ������</td>
        <td width="87%" class="category"><input type="text" name="truename" style="width:200px">
          &nbsp;&nbsp;</td>
      </tr>	  
      <tr>
	    <td height="30"></td>
        <td class="category">
		 <input type="hidden" name="actiont" value="add">
		  <input type="submit" value=" ȷ������ " onClick="return check()" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" value=" ������д " class="button">		</td>
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
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form name="form2">
  <input type="hidden" name="form" value="form2">
  <input type="hidden" name="actiont" value="search">
  <tr>  
    <td width="13%" height="30"><a href="#" onClick="UserOperate()">�����û�</a></td>
	<td width="87%" align="right">
	  �����û���ѯ��
	    
	    <input type="text" name="keyword" size="20" value="">
	  <input type="submit" value=" ��ѯ " class="button">&nbsp;	</td>
  </tr>
</form>  
</table>
<div id="userlist" align="center"><img src="../skin/Images/ajaxing.gif">���ڼ����û��б�...</div>
<div id="officeuserlist" style="display:none" >
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
	 ���
	   <a href="?action=order&orderid=UserID&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=UserID&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
    <td width="8%" class="category">
	  �û�����	  </td>
  
	<td width="16%" class="category">
	 ����½ʱ��   <a href="?action=order&orderid=LastLoginTime&Orderby=0"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=LastLoginTime&Orderby=1"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
	 <td width="23%" class="category">
	 ��������</td>
	 <td width="8%" class="category">
	 �û�״̬</td>
	<td width="13%" class="category">
	 ����Ȩ��</td>

    <td class="category" width="9%">�޸�</td>

	
    <td class="category" width="5%">ɾ��</td>
	
	
	<td class="category" width="8%">�鿴��ϸ</td>
  </tr>
  <% 
 Dim Ds,i,Count,Page,Pagesize,Cmd,FieldName(0),FieldValue(0),MaxPage,keyword,keystr,OrderBy,OrderField
 
 keyword = Jcs.Checkstr(Request("keyword"))
keystr = "1 = 1 And  (TrueName Like  '%"&keyword&"%' Or UserName Like  '%"&keyword&"%' )"
'-----
OrderBy = 1
OrderField = "UserID"
If Request.QueryString("action") = "order" Then 
If Not IsNull(Request.QueryString("orderid")) Then OrderField = Jcs.Checkstr(Request.QueryString("orderid"))

If Not IsNull(Request.QueryString("orderby")) Then OrderBy = Jcs.Checkstr(Request.QueryString("orderby"))
End If

Page = Request("Page")
If Page = "" Or IsNumeric(Page) = 0 Then Page = 1
Pagesize = 2
SQL = "Select Count(UserID) From [v_JcUser] Where 1 = 1  And "& keystr
Count = Jcs.Execute(SQL)(0)
If Count <= 0 Then
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>û���ҵ������û�</b></td></tr>"
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

Response.Write("<tr onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''"" onDblClick=""javascript:var win=window.open('User_View.Asp?id="&Ds(0,i)&"','�����û���ϸ��Ϣ','width=853,height=470,top=176,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()"" ><td align=""center"" height=""25"">"&Ds(0,i)&"</td><td align=""center"">"&Ds(1,i)&"</td><td align=""center"">"&Ds(3,i)&"</td><td align=""center"">"&Ds(4,i)&"</td><td align=""center"">"&LockStat(Ds(0,i))&"</td><td align=""center""><a href=""User_Level.Asp?id="& Ds(0,i)& """>�������Ȩ��</a></td>")
Response.Write("<td align=""center""><a href=""User_Edit.Asp?id="& Ds(0,i)& """><img src=""../skin/images/res.gif"" border=""0"" hspace=""2"" align=""absmiddle"">�޸�</a></td>")
Response.Write("<td align=""center""><input type=""checkbox"" name=""ID"" value=" &Ds(0,i)& " style=""border:0""></td><td align=""center""><a href=""User_View.Asp?id="& Ds(0,i)& """>����鿴</a></td></tr>")

 Next:Response.Flush

' Closedatabase()   '�ر����ݿ�����
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">&nbsp;
	  ˫��ÿ�пɲ鿴�û���ϸ����; ״̬�����"����"����û���������
	  </td>  
	<td width="31%" align="right">
	  <%Response.write(Jcs.PageList( Page,Pagesize,Count,FieldName,FieldValue) )%>	  &nbsp;</td>
	 
    <td width="20%" align="right"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="border:0"> <input type="hidden" name="page" value="<%=Request("Page")%>">
      ȫѡ <input type="hidden" name="actiondel" value="sub">
        <input name="submit" type="submit" class="button" onClick="return confirm('�˲����޷��ָ������������أ�����\n\nȷ��Ҫɾ����ѡ��Ļ�Ա��')" value="ɾ ��"></td>
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
</div>
<script language="javascript" type="text/javascript">
    setTimeout("setviewdata('userlist','officeuserlist')",1000);
</script>
</body>
</html>
<% 
End Sub 

'----------ɾ�������û�
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
'---д������־
Jcs_Manger.GetUser ID(i)
Username = Jcs_Manger.UserTitle
Jcs.WriteLog Username,"ɾ�������û�"
Jcs_Manger.ClearUser
Next 

'Closedatabase()   '�ر����ݿ�����
If J > 0 Then 
 Msg = "ɾ���ɹ���"
 Else 
 Msg = "����Ҫѡ��һ��Ҫɾ������"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 
'-------------------���ӱ����û�
Sub AddUser()
Dim Uname,Pwd,Tname,Tbofficeid,ND
Uname = Jcs.Checkstr(Request.Form("username"))
Pwd = Jcs.Checkstr(Request.Form("password"))
Tname = Jcs.Checkstr(Request.Form("truename"))
Tbofficeid = Jcs.Checkstr(Request.Form("officename"))
If Uname = ""   Then
	Founderr = True
	Msg = "�û�������Ϊ�գ�"
	Jcs.BackInfo "",0
	Exit Sub
End If
'--------------------�ж��û��Ƿ����
If Not Founderr Then
	SQL = "Select UserName From [Jc_User] Where UserName = '"&Uname&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Founderr = True
	Rs.Close:Set Rs = Nothing     '�ͷ�Rs
		Msg = "�û����Ѿ����ڣ�"
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
'---д������־
Jcs.WriteLog Tname,"���ӱ����û�"
	Closedatabase()   '�ر����ݿ�����
Msg = "�����û���" &Tname& "�����ӳɹ���"
		Jcs.BackInfo "OfficeUser_List.Asp",2
End If 
End Sub 

'------------------�ж�����״̬
Function LockStat(ByVal id)
Dim LockStr
 Jcs_Manger.GetUserInfo(id)
LockStr = Jcs_Manger.UDeleteFlag

Select Case LockStr
Case 1
LockStat = "<a href=?actiont=lock >��ɾ��</a>"
Case 2
LockStat = "<a href=?actiont=lock&ID="& id &"&page="&Request("Page")&" title=����ָ�ʹ��>������</a>"
Case Else 
LockStat = "<a href=?actiont=lock&ID="& id &"&page="&Request("Page")&" title=�������>����</a>"
End Select 
End Function
'-------------------�����û���ͨ�û�
Sub LockUserInfo()
Dim id,delflag,mothed,truename,page
page = Jcs.Checkstr(Request("page"))
id =  Jcs.Checkstr(Request("ID"))
If IsNull(id)  Then 
 Msg = "��������"
 Founderr = True
Jcs.BackInfo "",0
End If 

Jcs_Manger.GetUserInfo id
delflag = Jcs_Manger.UDeleteFlag
Select Case delflag
Case 0
SQL = "Update [Jc_User] Set DeleteFlag =  2 Where UserID = " & id
mothed = "�����û�"
Case 2
SQL = "Update [Jc_User] Set DeleteFlag =  0 Where UserID = " & id
mothed = "����û�����"
Case Else
Founderr = True
End Select 
If Not Founderr Then 
Jcs.Execute(SQL)
'---д������־
truename = Jcs_Manger.TrueName
Jcs.WriteLog truename,mothed
Jcs_Manger.ClearUserInfo
 Msg = mothed &"��" &truename&"���ɹ���" 
Else 
 Msg = "���û��Ѿ���ɾ��"
End If 
  Jcs.BackInfo "?page="&page,2
End Sub 
 %>
 