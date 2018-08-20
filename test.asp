<!-- #include file="inc/Const.Asp" -->

<%
Function GetKeyWord() 
Dim Ds1,str,i
Sql="Select KeyWord From Tb_KeyWord"
Set Rs=Jcs.Execute(Sql)
If Rs.Bof And Rs.Eof Then
	Ds1=Null
	GetKeyWord=Ds1
Else
	Ds1 = Rs.GetRows()
	For i=0 to Ubound(Ds1,2)
	If i=Ubound(Ds1,2) Then
	str=str&Ds1(0,i)
	Else
	str=str&Ds1(0,i)&"|"
	End If
	Next
	GetKeyWord=str
End If
Rs.Close:Set Rs = Nothing
End Function
response.write GetKeyWord()&"<p>"
dim ds2,k
ds2=Split(GetKeyWord(),"|")
For k=0 to ubound(ds2,1)
 response.write k&ds2(k)&"<br>"
Next
  
%>