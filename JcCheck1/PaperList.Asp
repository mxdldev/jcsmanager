<!-- #include file="../Inc/Const.Asp" -->
<% 'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'�жϵ�½
Jcs.CheckManger()
'----------���嵼��
Dim action,Navigation,PaperName,PublishDate
action = LCase(Request("action"))
PaperName = Jcs.Checkstr(Request("PaperName"))
PublishDate=Jcs.Checkstr(request("PublishDate"))
If action="search" or PaperName<>"" or PublishDate<>"" Then
Navigation = SystemName & ">> ��ֽ�������" &">>��ֽ��ѯ"
Else
Navigation = SystemName & ">> ��ֽ�������" &">>��ֽ����"
End If

Select Case action
Case "del"      'ɾ��
Call DelPaper()
Case Else 
Call MainSub()
End Select 
'-------------------------�����б������岿�ֿ�ʼ
Sub MainSub()
 Dim Ds,i,Count,Page,Pagesize,Cmd,MaxPage,keystr,OrderBy,OrderField,PaperDir,Check,stated,FieldName(1),FieldValue(1)
 %>

<html>
<head>
<title>��ֽ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../skin/css/style.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
	<script language=javascript>
	function OpenWin(str){
		var win_width=window.screen.availwidth-10;
		var win_height=window.screen.availheight-25;
		window.open (str,"", "fullscreen=1,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbar=yes,resizable=no,copyhistory=yes,width="+win_width+",height="+win_height+",top=0,left=0");
	}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
//ɾ����ʾ����
function ConfirmDel()
{
	if (document.myform.Action.value=="Del")
	{
	   if(confirm("ȷ��Ҫɾ��ѡ�еı�ֽ��"))
		 return true;
	   else
		 return false;
	}	 
}
  function CheckInput()
  {
    var PaperName
	 PaperName=document.SearchForm.PaperName.value;
	// alert(PaperName);
	 if(PaperName=='')
	 {
	   alert('������������Ҫ������ֽ�Ĺؼ��֣�');
	   return false;
	 }
	 else
	 {
	   return true;
	 }
  }
</script>
<script src="../skin/js/calenda1.js" language="javascript"></script>

</HEAD>

<BODY>

<% 
'----������ԱȨ��
 ' Jcs_Manger.Manageckeck()

'ȡ�õ�ǰҳ��
%>
<table width="100%" border="0" cellpadding="0" cellspacing="2" align="center">
<form action=""  name="SearchForm" >
  <tr> 
    <td width="37%" height="30">&nbsp;</td>
	<td width="63%" align="left">
	  ��ֽ��ѯ(��ֽ����)��
	    <input  type="hidden" name="action" value="search"> 
	    <input type="text" name="PaperName" size="20" value="">
	  <input type="submit" value=" ��ѯ " class="button" onClick="return CheckInput()">&nbsp;�������ڣ�<span style="color:#FF0000;">
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
    <td width="5%">ѡ��</td>
    <td width="10%" height="30" class="category">���<a href="?action=order&orderid=PaperID&Orderby=0&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PaperID&Orderby=1&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a></td>
<td width="20%" class="category">��ֽ����</td><td width="20%" class="category">��ֽ����</td><td width="15%" class="category">����ʱ��<a href="?action=order&orderid=PublishDate&Orderby=0&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/up2.gif" border="0" hspace="2" align="absmiddle"></a><a href="?action=order&orderid=PublishDate&Orderby=1&PaperName=<%=PaperName%>&PublishDate=<%=PublishDate%>"><img src="../skin/images/down2.gif" border="0" hspace="2" align="absmiddle"></a> </td>
   <td width="10%" class="category">״̬</td>
    <td class="category" width="20%">����</td>
  </tr>
  <% 
If PublishDate="" Then
  FieldName(0)="PaperName"
  FieldValue(0)=PaperName
keystr = "Deleted = 0 And Tb_Check=0 And UserID="&Jcs.UserID&" And PaperName Like  '%"&PaperName&"%'"
Else
  FieldName(0)="PaperName"
  FieldValue(0)=PaperName
  FieldName(1)="PublishDate"
  FieldValue(1)=PublishDate
keystr = "Deleted = 0 And Tb_Check=0 And UserID="&Jcs.UserID&" And PublishDate='"&PublishDate&"' And PaperName Like  '%"&PaperName&"%'"
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
Response.Write "<tr align=""center"" onMouseOver=""this.className='highlight'"" onMouseOut=""this.className=''""><td colspan=""10"" height=""25"" align=""center"" style=""color:red""><b>��ʱû��ҪԤ��ı�ֽ</b></td></tr>"
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
PaperDir="http://10.17.255.142/paper/"&Ds(2,i)&"/"&Ds(4,i)&"/index.htm"
Response.Write "<tr align=""center""><td><input name=""ID"" type=""checkbox"" id=""ID"" value='"&Ds(0,i)&"' style=""border:0""></td><td>"&Ds(0,i)&"</td><td>"&Ds(1,i)&"</td><td>"&Ds(2,i)&"</td><td>"&Ds(3,i)&"</td><td style=""color:#f00"">"&vbcrlf
Check=Cint(Ds(5,i))
Select Case Check
Case 0
 stated="δ���"
Case 1
 stated="����ͨ��"
Case 2
 stated="����δ��"
Case 3
 stated="����ͨ��"
End Select
 response.write stated
 Check=0
  Response.Write "</td><td><a href=""JavaScript:void(0);"" OnClick=""OpenWin('"&PaperDir&"');"">�鿴</a>|<a href=""FirReview.asp?action=check&id="&Ds(0,i)&""">����</a>|<a href=""?action=del&id="&Ds(0,i)&""" onclick=""return ConfirmDel()"">ɾ��</a></td></tr>"
'Response.Write "<tr align=""center""><td>"&Ds(0,i)&"</td><td>"&Ds(1,i)&"</td><td>"&Ds(2,i)&"</td><td>"&Ds(3,i)&"</td><td><a href="&PaperDir&">�鿴</a>|ɾ��</td></tr>"
 Next:Response.Flush
' Closedatabase()   '�ر����ݿ�����
 %>
  <tr>
    <td colspan="10" height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="49%" align="left" style="color:#FF0000;">
ȫѡ��<input name="chkAll" type="checkbox" id="chkAll" value="" onClick="CheckAll(this.form)" style="border:0"> <input name="Action" type="hidden" id="Action" value="Del"><input name="submit" type="submit" class="button"  value="����ɾ��ѡ���ı�ֽ"></td>
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
Closedatabase()   '�ر����ݿ�����
If J > 0 Then 
 Msg = "ɾ���ɹ���"
 Else 
 Msg = "����Ҫѡ��һ��Ҫɾ������"
End If 
 Jcs.BackInfo "?page="&page,2
End Sub 
 %>