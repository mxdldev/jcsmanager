<!-- #include file="Inc/Const.Asp" -->
<% 
'判断登陆
Jcs.CheckManger()
 %>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META HTTP-EQUIV="Refresh" content="300">
<STYLE type=text/css>
A:link {
	FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A:active {
	FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A:visited {
	FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A:hover {
	FONT-SIZE: 12px; COLOR: #f29422; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A.blue:link {
	FONT-SIZE: 12px; COLOR: #ffffff; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A.blue:active {
	FONT-SIZE: 12px; COLOR: #73a2d6; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A.blue:visited {
	FONT-SIZE: 12px; COLOR: #ffffff; LINE-HEIGHT: 20px; TEXT-DECORATION: none
}
A.blue:hover {
	FONT-SIZE: 12px; COLOR: #73a2d6; LINE-HEIGHT: 20px; TEXT-DECORATION: underline
}
</STYLE>
<SCRIPT language=JavaScript>
				nereidFadeObjects = new Object();
				nereidFadeTimers = new Object();
				function nereidFade(object, destOp, rate, delta){
				if (!document.all)
				return
					if (object != "[object]"){ 
						setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
						return;
					}
					clearTimeout(nereidFadeTimers[object.sourceIndex]);
					diff = destOp-object.filters.alpha.opacity;
					direction = 1;
					if (object.filters.alpha.opacity > destOp){
						direction = -1;
					}
					delta=Math.min(direction*diff,delta);
					object.filters.alpha.opacity+=direction*delta;
					if (object.filters.alpha.opacity != destOp){
						nereidFadeObjects[object.sourceIndex]=object;
						nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
					}
				}
				</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 scroll=no marginheight="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgColor=#3a6592 style="color:#FFFFFF; font-size:12px;">
  <tr>
    <td height="28" width="214"><a href="http://www.joyhua.com/" target="_blank"><IMG src="skin/Images/Title.gif" width="168" height="28" border="0"></a></td>
    <td width="68" nowrap="nowrap">&nbsp;</td>
    <td width="326"> <%Response.write(JccTitleName) %> - 管理系统 </td>
    <td width="21">&nbsp;</td>
    <td width="*" align="right" nowrap="nowrap">
	  [<A class=blue href="#" onClick="showModalDialog('Skin/about.htm',window,'dialogHeight:260px;dialogWidth:360px;dialogleft:200px;help:no;status:no;scroll:no');">关于系统</A>]
	  [<A class=blue href="Index.Asp?action=exit" target="parent" onClick="return confirm('确定要退出吗？');">安全退出</A>]
	  &nbsp;&nbsp;
	</td>
  </tr>
</table>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 background=Skin/images/topnav_bg.jpg>
  <TR>
    <TD width="181" height=35>&nbsp;</TD>
	<TD valign="bottom" width="74">
	  <a href="produit/produit_add.asp" target="right">用户管理</a>	</TD>
	<%  
	%>
	<TD valign="bottom" width="72">
	  <a href="produit/produit.asp" target="right"></a>	<a href="produit/produit_add.asp" target="right">用户管理</a></TD>
		
  </TR></TABLE>
</BODY></HTML>
