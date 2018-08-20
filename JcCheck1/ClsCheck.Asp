<%
'**************************************************************
'-- File Name: ClsCheck.Asp 管理操作类
'-- Version: 1.0 Beta build 20090310
'-- Date: 2009-3-17
'-- Script Written By guoxiaolong
'**************************************************************

Class Cls_Check
                                   '报社属性
Public Sub DelPaper()
Dim ID,i,J,page,PaperName,PublishDate
page = Jcs.Checkstr(Request("page"))
J = 0
ID =  Jcs.Checkstr(Request("id"))
PaperName=Jcs.Checkstr(Request("PaperName"))
PublishDate=Jcs.Checkstr(Request("PublishDate"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Tb_Paper] Set Deleted =  1 Where PaperID = " & ID(i)
Jcs.Execute(SQL)
Jcs.WriteLog GetPaperName(ID(i))&"(出版日期："&GetPubLishDate(ID(i))&")","删除报纸"
J = J + 1
Next 
Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项"
End If 
 Jcs.BackInfo "?page="&page&"&PaperName="&PaperName&"&PublishDate="&PublishDate,2
End Sub 

Public Sub DelArticle()
Dim ID,i,J,page
PaperId=Jcs.Checkstr(Request("PaperId"))
page = Jcs.Checkstr(Request("page"))
Title=Jcs.Checkstr(Request("Title"))
PublishDate=Jcs.Checkstr(Request("PublishDate"))
J = 0
ID =  Jcs.Checkstr(Request("ID"))
ID = Split(ID,",")
For i = 0 To UBound(ID)
SQL = "Update [Tb_Article] Set Deleted =  1 Where ArticleID = " & ID(i)
Jcs.Execute(SQL)
Jcs.WriteLog GetArticleTitle(ID(i)),"删除文章"
J = J + 1
Next 
Closedatabase()   '关闭数据库链接
If J > 0 Then 
 Msg = "删除成功！"
 Else 
 Msg = "至少要选择一个要删除的项！"
End If 
 Jcs.BackInfo "?page="&page&"&PaperId="&PaperId,2
End Sub 

Public Function GetPaperName(PaperId)
  Dim Ds,Sql,Rs
  Sql="select PaperName From Tb_Paper Where PaperId="&PaperId&""
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.Getrows()
  If IsArray(Ds) Then GetPaperName=Ds(0,0)
  Rs.Close():Set Rs=Nothing
End Function

Public Function GetPubLishDate(PaperId)
  Dim Ds,Sql,Rs
  Sql="select PubLishDate From Tb_Paper Where PaperId="&PaperId&""
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.Getrows()
  If IsArray(Ds) Then GetPubLishDate=Ds(0,0)
  Rs.Close():Set Rs=Nothing
End Function

Public Function GetPaperType(PaperId)
  Dim Ds,Sql,Rs
  Sql="select PaperType From Tb_Paper Where PaperId="&PaperId&""
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.Getrows()
  If IsArray(Ds) Then GetPaperType=Ds(0,0)
  Rs.Close():Set Rs=Nothing
End Function

Public Function GetArticleTitle(ArticleId)
  Dim Ds,Sql,Rs
  Sql="select Title From Tb_Article Where ArticleId="&ArticleId&""
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.Getrows()
  If IsArray(Ds) Then GetArticleTitle=Ds(0,0)
  Rs.Close():Set Rs=Nothing
End Function

PUblic Function GetClassName(ID)
 Dim Ds,Sql,Rs
 Sql="select Tb_ClassName From Tb_Class Where Tb_ClassId="&ID&""
  Set Rs=Jcs.Execute(Sql)
  If Not Rs.eof Then Ds=Rs.Getrows()
  If IsArray(Ds) Then GetClassName=Ds(0,0)
  Rs.Close():Set Rs=Nothing
End Function

PUblic Sub ClassReview()
 Dim id,sql,stat
 id=Jcs.CheckStr(Request("id"))
 stat=Jcs.CheckStr(Request("stat"))
 sql="Update Tb_Class Set Tb_ClassState="&stat&" Where Tb_ClassId="&id
 'Response.write sql
 Jcs.Execute(sql)
 Jcs.WriteLog GetClassName(id),"修改报刊审核意见"
 Response.redirect("TbClassList.asp")
End Sub

End Class

%>