<!-- #include file="../Inc/Const.Asp" -->
<% 
'判断登陆
Jcs.CheckManger()
 %>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<SCRIPT language=javascript>
var itemp;
var tobj="";
itemp="";
function leftBgOver(obj){//菜单
	obj.style.background="url(/skin/images/left_bg02.gif) center no-repeat";
	//obj.style.position="center";
	//obj.style.repeat="no-repeat":
}

function leftBgOut(obj,sty){
	//alert( (sty)?"url(" + sty + ")":"");
if (tobj!="")
{
sty="images/left_bg01.gif";
obj.style.background= (sty)?"url(" + sty + ")":"";
}
else 
{
obj.style.background= (sty)?"url(" + sty + ")":"";
}
}
</script>
<script language="javascript">
function collapse(objid)
{
	var obj = document.getElementById(objid);
	collapseAll();
	obj.style.display = '';
}
function collapseAll()
{
	for (var i=1; i<11; i++)
	{
		var obj = document.getElementById('g_' + i.toString());
		if (obj) obj.style.display = 'none';
	}
}
function expandAll()
{
	for (var i=1; i<11; i++)
	{
		var obj = document.getElementById('g_' + i.toString());
		if (obj) obj.style.display = '';
	}
}
</script>
<link href="/skin/Css/Style.css" rel="stylesheet" type="text/css">
<style rel="stylesheet" type="text/css">
body {margin:0px 5px;}
img{border:none;}
.menuall{text-align:center;width:149px;background:#ffffff;}
.menuall div{margin:0px 0 5px 10px;text-align:left;}
.STYLE1 {
	color: #FFFFFF;
	font-weight: bold;
}
</style>
</head>
<body>
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
<table border="0" cellpadding="0" cellspacing="0" class="menuall">
<tr>
<td><img src="/Skin/images/left_top.gif" alt="" /></td>
</tr>
<tr>
<td>
<a href="javascript:expandAll()" target="_self"><img src="/Skin/images/extand.gif" alt="展开菜单" onMouseOver=nereidFade(this,100,10,5) style="FILTER:alpha(opacity=50)" onMouseOut=nereidFade(this,50,10,5) /></a>&nbsp;<a href="javascript:collapseAll()" target="_self"><img src="/Skin/images/collapse.gif" alt="收拢菜单" onMouseOver=nereidFade(this,100,10,5) style="FILTER:alpha(opacity=50)" onMouseOut=nereidFade(this,50,10,5) /></a></td>
</tr>
<tr>
<td onClick="collapse('g_1')" height="25" style="cursor:hand;" background="/skin/Images/menu_a.gif"><span class="STYLE1">报纸阅读</span></td>
</tr>
<tr>
<td id="g_1"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

    <tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Paper_Read.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>报纸在线阅读</td></tr></table></td>
    </tr>
   
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>
<tr>
<td onClick="collapse('g_2')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">审读管理</span></td>
</tr>
<tr>
<td id="g_2"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>
    <tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='produit/produit_move.asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>查看审读情况</td></tr></table></td>
    </tr>

	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>
<tr>
<td onClick="collapse('g_3')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">用户管理</span></td>
</tr>
<tr>
<td id="g_3"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>
    <tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='OfficeUser_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td> 报社用户管理</td></tr></table></td>
    </tr>
    <tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='CheckUser_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>审读员管理</td></tr></table></td>
    </tr>
   	   <tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='ManUser_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>管理员管理</td></tr></table></td>
    </tr>	
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>


<tr>
<td onClick="collapse('g_6')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">关键字管理</span></td>
</tr>
<tr>
<td id="g_6"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='KeyWord_Add.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>添加关键字</td></tr></table></td>
    </tr>
	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='KeyWord_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>管理关键字</td></tr></table></td>
    </tr>
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>

<tr>
<td onClick="collapse('g_7')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">公告管理</span></td>
</tr>
<tr>
<td id="g_7"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Notice_Add.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>发布公告</td></tr></table></td>
    </tr>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Notice_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>管理公告</td></tr></table></td>
    </tr>
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>

<tr>
<td onClick="collapse('g_10')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">分配任务</span></td>
</tr>
<tr>
<td id="g_10"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Give_TaskPaper.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>给审读员分配报纸</td></tr></table></td>
    </tr>
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>

<tr>
<td onClick="collapse('g_9')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">报社管理</span></td>
</tr>
<tr>
<td id="g_9"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Office_Add.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>添加报社</td></tr></table></td>
    </tr>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='Office_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>管理报社</td></tr></table></td>
    </tr>
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>

<tr>
<td onClick="collapse('g_8')" style="cursor:hand;" height="25" background="/skin/Images/menu_a.gif"><span class="STYLE1">日志管理</span></td>
</tr>
<tr>
<td id="g_8"><table width="100%" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tbody>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='OfficeLog_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>查看报社日志</td></tr></table></td>
    </tr>

	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='CheckLog_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>查看审读日志</td></tr></table></td>
    </tr>
	<tr>
      <td height="25" align="center" background="/Skin/images/left_bg01.gif" style="cursor:hand"  onclick="javascript:parent.right.location.href='MangerLog_List.Asp';"><table cellpadding="0" cellspacing="0" width="100%"><tr><td width="50">&nbsp;</td><td>管理员日志</td></tr></table></td>
    </tr>
	<tr><td height="5"></td></tr>
  </tbody>
</table></td>
</tr>
</table>
<div style="font-size:11px;font-family:Tahoma; color:#CEE6FA" align="center">Powered
by<a href="http://www.joyhua.com" target="_blank"> <b style="color:#CEE6FA">锦华科技</b></a></div>
<script>
expandAll();
</script>
</body></html>