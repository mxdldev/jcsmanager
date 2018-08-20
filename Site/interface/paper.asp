<!--#include file="../inc/func.asp" -->
<%set conn=server.CreateObject("ADODB.Connection")
  conn.open GetConnectionString
        '特别注意:存在N个ID文件,而且每个ID对应的文件夹下面包括N个TRS   要做循环处理 从info.dat文件中提出制作者
		'------1.根据以后缀为"id"文件名命名的文件夹(根目录)下面读取TRS文件,并存储到数据库
		'------2.后再删除ID文件及后缀为"~ID"文件		
		'第一步(1)
		 Dim fso1, fa, f11, fc1
         Set fso1 = CreateObject("Scripting.FileSystemObject")
         'dim strFolder: strFolder= "lnrb,lswb,ssshdb,bdcb"
         dim strFolder: strFolder= "tzrb,tzwb"
         dim arrFolder : arrFolder = split(strFolder,",",-1,1)
         max  = ubound(arrFolder)
         for p = 0 to max         
			Set fa = fso1.GetFolder(server.MapPath("../../"&arrFolder(p)&"/"))  
			folerNamePaper = arrFolder(p)
			Set fc1 = fa.Files
			For Each f11 in fc1
				extName1 = fso1.GetExtensionName(server.MapPath("../../"&arrFolder(p)&"/"&f11.name))
				if extName1 = "id" then   
					baseName = fso1.GetBaseName(server.MapPath("../../"&arrFolder(p)&"/"&f11.name)) 
					if baseName<>"" then                                     
						dim rsExists : set rsExists = server.CreateObject("adodb.recordset")
						rsExists.Open "select id from tb_paper where folder='" & folerNamePaper&"/"&baseName & "'",conn,1,1
						if not rsExists.EOF then				       
						  conn.Execute("delete from tb_article where paperId=" & cint(rsExists("id")) )	
										  conn.Execute("delete from tb_paper where folder='" & folerNamePaper&"/"&baseName & "'")			    
						end if		
						'------------------------------------------------------------------------------ 
							  dim fsoy : Set fsoy = CreateObject("Scripting.FileSystemObject")	
												  If (fsoy.FileExists(server.MapPath("../../" & folerNamePaper&"/"&baseName & "/info.dat"))) Then
							  Set tf = fsoy.OpenTextFile(server.MapPath("../../" & folerNamePaper&"/"&baseName & "/info.dat"), 1)
							  proUserName = tf.ReadLine '制作者
												  else
													  proUserName = ""
												  end if                                              
							  paperType=folerNamePaper '报纸类型ID(string)
							  strPublishDate = baseName
							  PublishDate = cdate(mid(strPublishDate,1,4)&"-"&mid(strPublishDate,5,2)&"-"&mid(strPublishDate,7,2)) '发布日期					      

							  if Trim(paperType) = "tzrb" then paperName="泰州日报"
								if Trim(paperType) = "tzwb" then paperName="泰州晚报"

                             								dim strTbl : strTbl = ""
								dim strPaper : strPaper = ""
								strTbl = strTbl & "insert into tb_paper(paperType,paperName,publishDate,folder,proUserName)"
								'strTbl = strTbl & " values('" & paperType & "','" & paperName & "','" & publishDate & "','" & folerNamePaper&"/"&baseName & "','" & proUserName & "')"
								strTbl = strTbl & " values('','" & paperName & "','" & publishDate & "','" & folerNamePaper&"/"&baseName & "','" & proUserName & "')"
								conn.Execute(strTbl)	
						'------------------------------------------------------------------------------			
						'循环找出此文件夹下面的所有的TRS文件并进行处理
						Set fax = fso1.GetFolder(server.MapPath("../../" &folerNamePaper&"/"&baseName & "/")) 
						Set fc1x = fax.Files
						For Each f11x in fc1x  '查出此文件夹下面所有的TRS	
						  extName1 = fso1.GetExtensionName(server.MapPath("../../" & folerNamePaper&"/"&baseName & "/"&f11x.name))					   
						  if extName1 = "trs" and len(f11x.name)=36 then		'后缀名是.trs并且文件名长度是36			
								dim filsPath : filsPath = server.MapPath("../../" & folerNamePaper&"/"&baseName & "/" &f11x.name)							
								Dim fso2,f2,ts,fileName
								Set fso2 = CreateObject("Scripting.FileSystemObject")     
								Set cnrs = fso2.OpenTextFile(filsPath,1)
								'读取文件结束,处理并存储到数据库							
								dim zz : zz = 0	
								dim yy : yy = 0
									content =""	
								While Not cnrs.AtEndOfStream                                                                
									rsline = cstr(cnrs.ReadLine)
									if rsline = "" then
										  zz = zz + 1
									  if zz mod 2 <> 0 then                                                                      
										  startFlag = true
									  else  
										  startFlag = false
									  end if
									elseif replace(replace(rsline,"<","+"),">","+") = replace(replace("<REC>","<","+"),">","+") then  '<REC>
									zz = 0
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<版名>=","<","+"),">","+") then  '<版名>=
									zz = 0								  
									verName = mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))								   
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<日期>=","<","+"),">","+") then  '<日期>=
									zz = 0								  
									strDate = mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))								   
    								elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<版次>=","<","+"),">","+") then  '<版次>=
									zz = 0								  
									verOrder = mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<引题>=","<","+"),">","+") then  '<引题>=
									zz = 0								  
									leadTitle= mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<标题>=","<","+"),">","+") then  '<标题>=
									zz = 0								  
									title =mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<副题>=","<","+"),">","+") then  '<副题>=
									zz = 0								  
									title1 =mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<图像>=","<","+"),">","+") then  '<图像>=
									zz = 0								  
									images =mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,7) = replace(replace("<全文坐标>=","<","+"),">","+") then  '<全文坐标>=
									zz = 0								  
									coordinate =mid(replace(replace(rsline,"<","+"),">","+"),8,len(replace(replace(rsline,"<","+"),">","+")))
									
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,5) = replace(replace("<作者>=","<","+"),">","+") then  '<作者>=
									zz = 0								  
									author =mid(replace(replace(rsline,"<","+"),">","+"),6,len(replace(replace(rsline,"<","+"),">","+")))
									
									elseif mid(replace(replace(rsline,"<","+"),">","+"),1,6) = replace(replace("<正文>=","<","+"),">","+") then  '<正文>=
									zz = 0								  
									content = content & replace(rsline,"<正文>=","")
									else
									zz = 0
									content = content & rsline & "<br>"
									end if
    								
									if startFlag = true then				    					    
										dim rsPaper : set rsPaper = server.CreateObject("adodb.recordset")
										rsPaper.Open "select top 1 id from tb_paper order by id desc",conn,1,1
										if not rsPaper.EOF then
										  paperid=rsPaper("id") '杂志id
										else
										  paperid = -1  '表示不存在
										end if	
										content = replace(content,"<REC>","")
										content = EscapeString(content)
										strPaper = strPaper & "insert into tb_article(paperId,verName,publishDate,verOrder,leadTitle,title,title1,images,coordinate,author,content,trsFilePath,orderId)"
										strPaper = strPaper & " values (" & paperId & ",'" & verName & "','" & publishDate & "','" & verOrder & "','" & leadTitle & "','"& title & "','" & title1 & "','"&images&"','" & coordinate & "','"&author&"','" & content & "','" & folerNamePaper&"/"&baseName&"/"&f11x.name & "'," & yy & ")"              						

										conn.Execute(strPaper)
										yy =yy +1
										startFlag=false		
										content=""	
										strPaper= ""
									end if								
								Wend						
						  end if
						next
							  '第二步(2) 删除id文件
								Set fsok = CreateObject("Scripting.FileSystemObject")
								If (fsok.FileExists(server.MapPath("../../" & folerNamePaper&"/"&baseName & ".id"))) Then
									fsok.DeleteFile(server.MapPath("../../" & folerNamePaper&"/"&baseName& ".id"))
								End If
								If (fsok.FileExists(server.MapPath("../../" & folerNamePaper&"/"&baseName& ".~id"))) Then
									fsok.DeleteFile(server.MapPath("../../" & folerNamePaper&"/"&baseName& ".~id"))
								End If
    					        
					end if
				end if
			Next
		next
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../css/css.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" bottommargin="0" rightmargin="0" leftmargin="0" oncontextmenu = "return false" onselectstart="return false">
<table width="100%"  border="0" cellspacing="0" cellpadding="0" ID="Table2">
  <tr>
    <td align="center" height="60" valign="middle"><font color="#ff0066">报纸数据已成功导入数据库！</font></td>
  </tr>
</table>
</body>
</html>





