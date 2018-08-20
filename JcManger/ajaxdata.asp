<!-- #include file="Const.Asp" -->
<%
Response.Charset = "GB2312" '不设置中文会乱码
'Response.write "<img src='../skin/Images/ajaxing2.gif'" width='10' height='10'>"
Dim ID,Result,Action
Action = Request("action")
ID = Request("ID")
Select Case Action
Case "user"
Call usercheck()
Case "keyword"
Call wordcheck()
Case "office"
Call officecheck()
End Select 
'----------------------
Sub usercheck()
If IsNull(ID) or ID = "" Then 
ResPonse.write("")
Else 
SQL = "Select UserName From [Jc_User] Where UserName = '"&ID&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Response.write("1") '用户名已经存在
	Rs.Close:Set Rs = Nothing     '释放Rs
	Else 
Response.write("0") '用户名可以使用
	End If 
End If 
End Sub 
'----------------关键字
Sub wordcheck()
If IsNull(ID) or ID = "" Then 
ResPonse.write("")
Else 
SQL = "Select KeyWord From [Tb_KeyWord] Where KeyWord  = '"&ID&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Response.write("1") '关键字已经存在
	Rs.Close:Set Rs = Nothing     '释放Rs
	Else 
Response.write("0") '关键字可以使用
	End If 
End If 
End Sub 
'----------------报社名
Sub officecheck()
If IsNull(ID) or ID = "" Then 
ResPonse.write("")
Else 
SQL = "Select Tb_OfficeName From [Tb_Office] Where Tb_OfficeName  = '"&ID&"'  "
	Set Rs = Jcs.Execute(SQL)
	If Not  Rs.Bof  Then
	Response.write("1") '报社名已经存在
	Rs.Close:Set Rs = Nothing     '释放Rs
	Else 
Response.write("0") '报社名可以使用
	End If 
End If 
End Sub 
	
%>
