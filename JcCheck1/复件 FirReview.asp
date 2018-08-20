<!-- #include file="../Inc/Const.Asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>报纸初审</title>
</head>

<body>
<%
Jcs.CheckManger()
 Dim Id,Action,KeyData,Ds
  Action=Jcs.Checkstr(Request("action"))
  id=Jcs.Checkstr(Request("id"))
Select Case Action
Case "check"
Call CheckPaper(id)
Case Else
Msg = "错误的来源！"
Jcs.BackInfo "/Desk.asp",2
'ChkBadWords("靠你，eeee，暗示发疯歌曲问题前头痛")
End Select

Function GetKeyWord() 
Dim Ds1
Sql="Select KeyWord From Tb_KeyWord"
Set Rs=Jcs.Execute(Sql)
If Rs.Bof And Rs.Eof Then
	Ds1=Null
	GetKeyWord=Ds1
Else
	Ds1 = Rs.GetRows()
	GetKeyWord=Ds1(0,0)
End If
Rs.Close:Set Rs = Nothing
End Function

Function ChkBadWords(Str)
     Dim KeyData,BadWords,RedWord
	 KeyData=Null
	 BadWords=Null
        KeyData=Split(GetKeyWord(),"|")
		If IsNull(Str) Or IsNull(KeyData) Then Exit Function
		Dim i
		For i = 0 To UBound(KeyData)
		   'Response.Write KeyData(i)&"<br>"
			If InStr(Str,KeyData(i))>0 Then
			 ' response.Write "-----------------------------------------------------"
			  RedWord="<font style=""font-size:14px;color:#f00"">"&KeyData(i)&"</font>"
	          Str= Replace(Str,KeyData(i),RedWord)
	          BadWords=BadWords&KeyData(i)&"##"
	        '  Response.Write str
			End If
		Next
		If IsNull(BadWords) Then
		   ChkBadWords=Null
		Else
		   ChkBadWords = Str&"@@@"&BadWords
		End IF
End Function

Sub CheckPaper(id)
Dim j,i,result,Sql1,totalnumber,CurrentPage,d,sType
  j=0

If Not IsEmpty(Request("p")) And Trim(Request("p")) <> "" Then
	CurrentPage = CLng(Request("p"))
Else
	CurrentPage = 0
End If
If Cint(Request("Nums"))="" or Cint(Request("Nums"))=0 Then
  Sql1="Select Count(ArticleID) From Tb_Article Where PaperId="&id&""
  totalnumber=Cint(Jcs.Execute(Sql1)(0))
  If totalnumber = "" Or totalnumber = 0 Then
	Msg = "该报纸下没有文章内容"
	Jcs.BackInfo "/Desk.asp",2
	Exit Sub
   End If
Else
totalnumber =Cint(Request("Nums"))
End IF
'Response.Write totalnumber
'Response.End()

Response.Write "<br><br>" & vbNewLine
Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
Response.Write "  <tr>" & vbNewLine
Response.Write "    <td height='50'>本次共需审查 <font color='blue'><b>" & totalnumber & "</b></font> 条文章信息<a id=txt1></a></td>" & vbNewLine
Response.Write "  </tr>" & vbNewLine
Response.Write "      <tr>" & vbNewLine
Response.Write "        <td ><table width='" & Fix((CurrentPage / totalnumber) * 400) & "'  border='0' cellpadding='0' cellspacing='0' style=""background: #7c86ea;height:20px""><tr><td align='center' style='font-size:12px'>" & FormatNumber(CurrentPage / totalnumber * 100, 2, -1) & " %</td></tr></table></td>" & vbNewLine
Response.Write "      </tr>" & vbNewLine
Response.Write "  <tr>" & vbNewLine
Response.Write "    <td align='center' id=txt2></td>" & vbNewLine
Response.Write "  </tr>" & vbNewLine
Response.Write "</table>" & vbNewLine
Response.Flush

If CurrentPage >= totalnumber Then
	Response.Write "<table width='400' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbNewLine
	Response.Write "<tr><td height='30' align='center'><input type='button' name='stop' value=' 审查完毕 ' onclick=""window.location.href='SecPaperList.Asp';"" class=button></td></tr>" & vbNewLine
	Response.Write "</table>" & vbNewLine
	Response.Flush
	Response.End
End If
 
  Sql="Select ArticleID,PaperId,Content From Tb_Article Where PaperId="&id&" Order by ArticleId Desc"
 ' response.Write Sql
  'response.End()
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.GetRows()
  Rs.Close:Set Rs = Nothing
 
  If IsArray(Ds) Then
    For i=0 to Ubound(Ds,2)
	 
	   If Not IsNull(ChkBadWords(Ds(2,i))) Then
	   'Response.Write ChkBadWords(Ds(2,i))
	      result=Split(ChkBadWords(Ds(2,i)),"@@@")
	      If IsArray(result) Then
	        ArtResult Ds(0,i),result(0),result(1)
		    j=j+1
	      End If
		End If
	Response.Write "<script>"
	Response.Write "txt1.innerHTML='，已成功审核：<B>" & i+1 & "</B> 条';"
	Response.Write "txt2.innerHTML='文章ID号：<B>" &Ds(0,i)& "</B> ';"
	Response.Write "</script>" & vbCrLf
	Response.Flush
	
Response.Write "<script language='JavaScript'>" & vbNewLine
Response.Write "function buildRefresh(){window.location.href='FirReview.Asp?action=check&id="&id&"&p=" & CurrentPage + 1 & "&Nums="&totalnumber&"';}" & vbNewLine
Response.Write "setTimeout('buildRefresh()',5);" & vbNewLine
Response.Write "</script>" & vbNewLine
Response.Flush

	Next
  End If
  
    If j>0 Then
     ParResult id,2
    Else
     ParResult id,1
    End If
End Sub

Sub ArtResult(id,Str,BadWords)
    If Trim(id)="" Then
	Msg = "错误的来源！"
	Jcs.BackInfo "/Desk.asp",2
    End If
  Sql="Update Tb_Article Set Content='"&Str&"',BadWords='"&BadWords&"' Where ArticleId="&id&""
  Jcs.Execute(Sql)
End Sub

Sub ParResult(id,result)
    If Trim(id)="" Then
	Msg = "错误的来源！"
	Jcs.BackInfo "/Desk.asp",2
    End If
  Sql="Update Tb_Paper Set Tb_Check="&result&" Where PaperId="&id&""
  Jcs.Execute(Sql)
End Sub
%>
</body>
</html>
