<!-- #include file="Const.Asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报纸预审</title>
<link href="../Skin/css/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
Jcs.CheckManger()
 Dim Id,Action,Ds,Navigation
  Action=Lcase(Jcs.Checkstr(Request("action")))
  id=Jcs.Checkstr(Request("id"))
  Navigation = SystemName & ">> 报纸预审管理" &">>文章审核"
Select Case Action
Case "show"
Call ShowArt(id)
Case "addreview"
Call AddReview()
Case Else
Msg = "错误的来源！"
Jcs.BackInfo "/Desk.asp",2
End Select

Sub ShowArt(id)
  
  Dim Sql,Ds,Rs1,Sql1,Ds1,ReviewContent
  
  Sql="select Title,PaperId,PublishDate,Images,ImageChecks From Tb_Article Where ArticleId="&id
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.GetRows()
  If IsArray(Ds) Then
Response.Write "<form name=""form1"">"
Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#C4D8ED"">"&vbcrlf
Response.Write "<tr>"&vbcrlf
Response.Write "<td><img src=""../skin/images/r_1.gif"" alt="""" /></td>"&vbcrlf
Response.Write "<td width=""100%"" background=""../skin/images/r_0.gif"">"&vbcrlf
Response.Write "<table cellpadding=""0"" cellspacing=""0"" width=""100%"">"&vbcrlf
Response.Write "<tr><td>&nbsp;"&Navigation&"</td><td align=""right"">&nbsp;</td></tr>"&vbcrlf
Response.Write "</table>"&vbcrlf
Response.Write "</td>"&vbcrlf
Response.Write "<td><img src=""../skin/images/r_2.gif"" alt="""" /></td>"&vbcrlf
Response.Write "</tr>"

Response.Write "<tr>"&vbcrlf
Response.Write "<td></td>"&vbcrlf
Response.Write "<td>"&vbcrlf
Response.Write "<table align=""center"" cellpadding=""4"" cellspacing=""1"" class=""toptable grid"" border=""1"">"&vbcrlf
Response.Write "<tr>"&vbcrlf
Response.Write "<td width=""25%"" align=""right"">文章标题：</td>"&vbcrlf
Response.Write "<td width=""75%"" class=""category1"">"&Ds(0,0)&"</td>"&vbcrlf
Response.Write "</tr>"&vbcrlf  
Response.Write "<tr>"&vbcrlf
Response.Write "<td align=""right"">出版时间：</td>"&vbcrlf
Response.Write "<td class=""category1"">"&Ds(2,0)&"</td>"&vbcrlf
Response.Write "</tr>"&vbcrlf  
Response.Write "<tr>"&vbcrlf
Response.Write "<td align=""right"">相关图片：</td>"&vbcrlf
Response.Write "<td class=""category1"">"
If Trim(Ds(3,i))<>"" or Ds(3,i)<>Null Then
  Dim arrpic,i,k,picurl
  k=0
  Response.Write "<table width=""100%"">"
  Response.Write "<tr>"
  arrpic=split(Ds(3,i),";")
  For i=0 to Ubound(arrpic,1)
  picurl="http://"&Request.ServerVariables("HTTP_HOST")&"/paper/"&Jcs_Check.GetPaperType(Ds(1,0))&"/"&Jcs_Check.GetPaperType(Ds(1,0))&"/"&Jcs.ShowTime(Ds(2,0),3)&"/m_"&arrpic(i)
    Response.Write "<td><img src="&picurl&" width=""200"" height=""200""><td>"
	k=k+1
	If k>2 Then
	  Response.Write "</tr><tr>"
	  k=0
	End If
  Next
  Response.Write "</tr>"
  Response.Write "</table>"
End If
Response.Write "</td>"&vbcrlf
Response.Write "</tr>"&vbcrlf 
Response.Write "<tr>"&vbcrlf
Sql1="select ReportContent From Check_Report Where ReportClass=2 And ArticleID="&id
Set Rs1=Jcs.Execute(Sql1)
If Not Rs1.eof Then 
Ds1=Rs1.Getrows()
ReviewContent=Ds1(0,0)
Rs1.Close():Set Rs1=Nothing
End If
Response.Write "<tr>"&vbcrlf
Response.Write "<td align=""right"">审核：</td>"&vbcrlf
Response.Write "<td class=""category1""> 通过：<input type=""radio"" name=""artcheck"" value=1 style=""border:none"""
If Ds(4,0)=1 Then Response.Write "checked"
Response.Write ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;不通过<input type=""radio"" name=""artcheck"" value=2 style=""border:none"""
If Ds(4,0)=2 Then Response.Write "checked"
Response.Write "></td>"&vbcrlf
Response.Write "</tr>"&vbcrlf 
Response.Write "<tr>"&vbcrlf
Response.Write "<td width=""25%"" height=""30"" align=""right"">审核意见：</td>"&vbcrlf
Response.Write "<td width=""75%"" class=""category""><textarea name=""ReviewContent"" style=""display:none"">"&ReviewContent&"</textarea><iframe ID=""Editor"" name=""Editor"" src=""../JcsEditor/index.html?ID=ReviewContent"" frameBorder=""0"" marginHeight=""0"" marginWidth=""0"" scrolling=""No"" style=""height:248px;width:550px""></iframe></td>"&vbcrlf
Response.Write "</tr>"&vbcrlf	
Response.Write "<tr>"&vbcrlf
Response.Write "<td height=""30"">&nbsp;</td>"&vbcrlf
Response.Write "<td class=""category"">"&vbcrlf
Response.Write "<input type=""hidden"" name=""action"" value=""AddReview"">"&vbcrlf
Response.Write "<input type=""hidden"" name=""id"" value="&id&">"&vbcrlf
Response.Write "<input type=""hidden"" name=""PaperId"" value="&Ds(1,0)&">"&vbcrlf
Response.Write "<input type=""submit"" value="" 确认添加 "" class=""button"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""reset"" value="" 重新填写 "" class=""button"">"&vbcrlf
Response.Write "</td>"&vbcrlf
Response.Write "</tr>"&vbcrlf
Response.Write "</table>"&vbcrlf
Response.Write "</td>"&vbcrlf
Response.Write "<td></td>"&vbcrlf
Response.Write "</tr>"&vbcrlf

Response.Write "<tr>"&vbcrlf
Response.Write "<td><img src=""../skin/images/r_4.gif"" alt="""" /></td>"&vbcrlf
Response.Write "<td></td>"&vbcrlf
Response.Write "<td><img src=""../skin/images/r_3.gif"" alt="""" /></td>"&vbcrlf
Response.Write "</tr>"&vbcrlf

Response.Write "</table>"&vbcrlf
Response.Write "</form>"
End If
End Sub

Sub AddReview()
  Dim Content,id,Sql,Ds,artcheck,PaperId
  id=Jcs.Checkstr(Request("id"))
  Content=Jcs.CheckStr(Request("ReviewContent"))
  artcheck=Jcs.Checkstr(Request("artcheck"))
  PaperId=Jcs.Checkstr(Request("PaperId"))
  If artcheck="" Then
    Msg = "请选择审核选项！"
	Jcs.BackInfo "SecReview.asp?action=show&id="&id&"",2
	Exit Sub
  End If
  If Content="" Then
    Msg = "审核意见不能为空！"
	'Jcs.BackInfo "SecReview.asp?action=show&id="&id&"",2
	Jcs.BackInfo "SecReview.asp",0
    Exit Sub
  Else
    Sql="Select RID From Check_Report Where ArticleId="&id
	Set Rs=Jcs.Execute(Sql)
	If Not Rs.eof Then
	 Ds=Rs.GetRows()
	 Sql="Update Check_Report Set ArticleID="&id&",ReportContent='"&Content&"',ReportClass=2,UserID="&Jcs.UserId&" Where RId="&Ds(0,0)
	 Jcs.WriteLog Jcs_Check.GetArticleTitle(id),"修改图片审核意见"
	 Else
    Sql="insert into Check_Report (ArticleID,ReportContent,ReportClass,UserID) values("&id&",'"&Content&"',2,"&Jcs.UserId&")"
	Jcs.WriteLog Jcs_Check.GetArticleTitle(id),"添加图片审核意见"
	End If
	'response.Write sql
	'response.end
    Rs.Close():Set Rs=Nothing
	Jcs.Execute(Sql)
	Msg="审核意见添加成功！"
      Sql="Update Tb_Article Set ImageChecks="&artcheck&" Where ArticleId="&id
	  Jcs.Execute(Sql)
	'Jcs.BackInfo "PicArtList.asp?PaperId="&PaperId&"",2
	Jcs.BackInfo "PicArtList.asp",2
  End If
End Sub

%>

</body>
</html>
