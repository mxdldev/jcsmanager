<%
'**************************************************************
'-- File Name: Cls_Manger.Asp �����û�������
'-- Version: 1.0 Beta build 20090327
'-- Date: 2009-3-27
'-- Script Written By Jiangtao
'**************************************************************
Class Cls_Office
Public Tb_classID,Tb_ClassName,Tb_Number,Tb_code,Editor,Tb_Dir,TbOfficeID,Tb_Date   '��������
Public Tb_OfficeName,Proprieter,Editor_in_chief,Tb_Telephone,Manager,AddDate,Tb_Address        '��������
'---------------------��鱨���û�Ȩ��
Public  Function  Officeckeck()
If Jcs.UserClassID <> 3 Then 
Response.Write("<br><center><img src=""../skin/images/note.gif"" align=""absmiddle"">&nbsp;<font color=""#FF0000"">�㲻�߱���Ȩ�ޣ��������Ա��ϵ��</font></center>")
Response.end
End If
End Function 

' *************************************
' ��ѯ��������
' @param officeid ����ID
' 2009-3-23 By Sunfeng
' *************************************
Function FindOffice(ByVal officeid)
	'Dim officeid
	Dim SqlOffice,RsOffice,Tb_OfficeName
	If officeid <> "" Then
		SqlOffice = "Select Tb_OfficeName From Tb_Office Where Tb_OfficeID=" & Officeid
		Set RsOffice = Jcs.Execute(SqlOffice)
		Tb_OfficeName = RsOffice("Tb_OfficeName")
		RsOffice.Close:Set RsOffice = Nothing
	End If
	FindOffice = Tb_OfficeName
End Function 
' *************************************
' ��ȡ��������
' @param id �������
' 2009-4-7 By jiangtao
' *************************************
Function GetTbClass(ByVal id)
Dim Usr
If ID = "" Or IsNull(ID) Then
    Founderr = True
	Msg = "������ʧ��"
	Jcs.BackInfo "",0
	Exit Function 
End If

If Not Founderr Then
	SQL = "Select Tb_classID,Tb_ClassName,Tb_Number,Tb_code,Editor,Tb_Date,Tb_OfficeID From [Tb_Class] Where Tb_classID = "&ID
	Set Rs = Jcs.Execute(SQL)
	If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   Msg = "������Ϣ�����ڻ��Ѿ���ɾ����"
		Jcs.BackInfo "",0
    Else 
    Usr = Rs.GetRows()
    Tb_ClassName = Usr(1,0)
    Tb_Number =  Usr(2,0)
    Tb_code =  Usr(3,0)
    Editor = Usr(4,0)
    Tb_Date =  Usr(5,0)
    TbOfficeID =  Usr(6,0)
  
	Rs.Close:Set Rs = Nothing     '�ͷ�Rs
	Usr = Null 
	End If
	'Closedatabase()   '�ر����ݿ�����
End If 
End Function 
Rem ----------------------------------------��ȡ������Ϣ
Public Function GetOffice(ByVal ID)
Dim Usr
If ID = "" Or IsNull(ID) Then
    Founderr = True
	Msg = "������ʧ��"
	Jcs.BackInfo "",0
	Exit Function 
End If

If Not Founderr Then
	SQL = "Select Tb_OfficeID,Tb_OfficeName,Editor_in_chief,Proprieter,Tb_Telephone,Manager,AddDate,Tb_Address From [Tb_Office] Where Tb_OfficeID = "&ID
	Set Rs = Jcs.Execute(SQL)
	If   Rs.Bof Or Rs.Eof Then
	Founderr = True
   Msg = "������Ϣ�����ڻ��Ѿ���ɾ����"
		Jcs.BackInfo "",0
    Else 
    Usr = Rs.GetRows()
    Tb_OfficeName = Usr(1,0)
    Proprieter =  Usr(3,0)
    Editor_in_chief =  Usr(2,0)
    Tb_Telephone = Usr(4,0)
	Manager =  Usr(5,0)
    AddDate = Usr(6,0)
	Tb_Address = Usr(7,0)
	
	Rs.Close:Set Rs = Nothing     '�ͷ�Rs	
	Usr = Null 
	End If
	'Closedatabase()   '�ر����ݿ�����
End If 
End Function 
'----------------ͨ��userid��ȡ������
Public  Function GetOfficename(userid)
Dim Rs
If IsNull(userid) Then 
GetOfficename = ""
Exit Function 
End If 
SQL = "SELECT UofficeName FROM v_JcUser WHERE (UserID = "&userid&")"
Set Rs = Jcs.Execute(SQL)
If Not rs.eof Then GetOfficename = Rs(0)
Rs.Close:Set Rs = Nothing     '�ͷ�Rs
End Function 
'----------------ͨ��userid��ȡ����id
Public  Function GetOfficeid(userid)
Dim Rs
If IsNull(userid) Then 
GetOfficeid = ""
Exit Function 
End If 
SQL = "SELECT Tb_OfficeId FROM Jc_User WHERE (UserID = "&userid&")"
Set Rs = Jcs.Execute(SQL)
If Not rs.eof Then GetOfficeid = Rs(0)
Rs.Close:Set Rs = Nothing     '�ͷ�Rs
End Function 
'---------------------��ȡ������ϸ��Ϣ


Rem ----------------------------��ձ������б�����Ϣ
Public Function ClearTbcalss()
   Tb_ClassName = Null 
   Tb_Number = Null
   Tb_code = Null 
   Editor = Null 
   Tb_Date = Null 
   TbOfficeID =Null 
End Function 
Rem ----------------------------��ձ������б�����Ϣ
Public Function ClearOffice()
   Tb_OfficeName = Null 
   Proprieter = Null
   Editor_in_chief = Null 
   Tb_Telephone = Null 
   Manager = Null 
   AddDate =Null 
   Tb_Address = Null 
End Function 
' *************************************

End Class 
%>