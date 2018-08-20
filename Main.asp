<!-- #include file="Inc/Const.Asp" -->
<% '强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
'判断登陆
Jcs.CheckManger()
'------根据用户级别分别调用各自文件夹里的Tree.Asp文件
Dim userclassid,treefile,rightfile
userclassid=Jcs.UserClassID
If Not  IsNumeric(userclassid) Then 
 Response.Redirect "/Index.Asp?action=exit"
 Else
 Select Case userclassid
 Case 1
 treefile="JcManger/Tree.Asp"
 rightfile="JcManger/Desk_Maner.Asp"
 Case 2
treefile="JcCheck/Tree.Asp"
rightfile="JcCheck/Desk_Check.Asp"
Case Else 
treefile="JcOffice/Tree.Asp"
rightfile="JcOffice/Office_Desk.Asp"
End Select 
End If 

'------------框架部分
Dim htmlstr
htmlstr = htmlstr +  "<html>"
htmlstr = htmlstr +  "<head>"
htmlstr = htmlstr +  "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"
htmlstr = htmlstr +  "<title>"& JccTitleName &" - 管理系统</title>"
htmlstr = htmlstr +  "</head><frameset rows=""63,*,32"" cols=""*"" frameborder=""no"" border=""0"" framespacing=""0"">"
htmlstr = htmlstr +  "<frame src=""Top.Asp"" name=""topFrame"" scrolling=""No"" noresize=""noresize"" id=""topFrame"" title=""topFrame"" />"
htmlstr = htmlstr +  "<frameset rows=""*"" cols=""177,*"" framespacing=""0"" frameborder=""no"" border=""0"">"
htmlstr = htmlstr +  "<frame src=""" & treefile & """ name=""left"" scrolling=""yes"" noresize=""noresize"" />"
htmlstr = htmlstr +  "<frame src=""" & rightfile & """ name=""right"" scrolling=""yes"" noresize=""noresize"" />"
htmlstr = htmlstr +  "</frameset>"
htmlstr = htmlstr +  "<frame src=""bottom.asp"" name=""bottomFrame"" scrolling=""No"" noresize=""noresize"" id=""bottomFrame"" title=""bottomFrame"" />"
htmlstr = htmlstr +  "</frameset>"
htmlstr = htmlstr +  "<noframes><body>"
htmlstr = htmlstr +  "</body>"
htmlstr = htmlstr +  "</noframes>"
htmlstr = htmlstr +  "</html>"
Response.Write(htmlstr)

 %>
