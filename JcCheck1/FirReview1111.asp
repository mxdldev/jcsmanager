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
  Sql1="Select Count(ArticleID) From Tb_Article Where PaperId="&id&""
  Count=Jcs.Execute(Sql1)(0)
  Dim 
If Not IsEmpty(Request("p")) And Trim(Request("p")) <> "" Then
	CurrentPage = CLng(Request("p"))
	d = Request("d")
Else
	CurrentPage = 0
	d = Timer()
End If
totalnumber = Request("Nums")
If Not IsNumeric(totalnumber) Or totalnumber = "" Or totalnumber = "0" Then
	Msg = "生成数量不正确"
	BackInfo "",0
	Exit Sub
End If

  Sql="Select ArticleID,PaperId,Content From Tb_Article Where PaperId="&id&" Order by ArticleId Desc"
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.GetRows()
  Rs.Close:Set Rs = Nothing
 
  If IsArray(Ds) Then
    For i=0 to Ubound(Ds,2)
	 
	   If Not IsNull(ChkBadWords(Ds(2,i))) Then
	   Response.Write ChkBadWords(Ds(2,i))
	      result=Split(ChkBadWords(Ds(2,i)),"@@@")
	      If IsArray(result) Then
	        ArtResult Ds(0,i),result(0),result(1)
		    j=j+1
	      End If
		End If
		
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
