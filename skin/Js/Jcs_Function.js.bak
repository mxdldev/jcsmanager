// 洪江涛 2009/03/31 js函数
function ValidatorTrim(s) { 
    var m = s.match(/^\s*(\S+(\s+\S+)*)\s*$/); 
    return (m == null) ? "" : m[1]; 
} 
// 是否为正整数
function isUnsignedInteger(strInteger){  
    var  newPar=/^\d+$/  
    alert(newPar.test(strInteger));
} 
// 是否为整数
function isInteger(strInteger){  
    // /^(-|\+)?\d+$/ 
    var  newPar=/^(- |\+)?\d+$/;
    return newPar.test(strInteger); 
}  
//用户名规则验证
//密码规则验证

// Grid 控制
function openWinows(url)
{
    var objWindow = window.open(url);
    if(objWindow){ }else{window.location.href = url;}
}
// 获取选中的 CheckBox 参数值
function GetChkItems(){ 
    var Items = new Array();
    for(var i=0;i<document.all("ItemChk").length;i++)
    { 
        if(document.all("ItemChk")(i).checked) 
        Items[Items.length] = document.all("ItemChk")(i).id; 
    }
    return Items;
}
// CheckBox 全选或全部取消     SetCheck
function SelectAll(){  // if(document.all[i].name=="ITEMTR") itemsi chkParams \\var m = document.getElementById("K").value;
    for(i=0;i<document.all("ItemChk").length;i++)
    {
        if(document.all("itemsi").checked)
        {
            document.all("ItemChk")(i).checked=true;
        }
        else
        {
            document.all("ItemChk")(i).checked=false;
        }
    }
}
// 使CheckBox只支持单选
function SetCheckBoxClear(objID){  
    var cbx = document.all("ItemChk"); 
    for(var i=0; i<cbx.length; i++)     
    {         
        if(cbx[i].type=="checkbox" && cbx[i].value == objID)         
        {             
            cbx[i].checked = true; 
        }
        else
        {
            cbx[i].checked = false; 
        }
    } 
    //objCheckBox.checked = true; 
}
// 改变背景颜色
function ChangeBGColor(ObjGrid)
{
    var Items = new Array();
    for(var i=0;i<document.all("YGVtr").length;i++){ 
    document.all("YGVtr")(i).style.backgroundColor='';document.all("YGVtr")(i).style.color='';}
    ObjGrid.style.backgroundColor='#336699';ObjGrid.style.color='#ffffff';
    //ObjGrid.style.backgroundColor='#99ccff';ObjGrid.style.color='#ffffff';
}
// 转到操作对象
function ChangeUrl(objUrl,objName)
{
    if(objUrl.length > 0)
    {
        if(objUrl.indexOf("Url=true")>1)
        {
            window.location.href = objUrl;
        }
        else
        {
            var objID = document.getElementById('selectRowID').value;
            var objFuncNo = document.getElementById('tbFuncNo').value;
            var objParams = document.getElementById('selectRowPara').value;
            if(CheckChangeParams(objFuncNo,objParams)>0)
            {
                if(isInteger(objID))
                {
                    window.location.href = objUrl + "&k="+objID+"&oNa="+objName;
                }
                else
                {
                    alert("操作失败：请首先从列表中点击选择您想要操作的项目！");
                }
            }
        }
        
    }
}
// 转到操作对象，包含选定的多行
function ChangeUrlMultCheck(objUrl,objName)
{
    if(objUrl.length > 0)
    {
        if(objUrl.indexOf("Url=true")>1)
        {
            window.location.href = objUrl;
        }
        else
        {
            var objArrayIDs = new Array();
            var selectIDs = "";
            var objFuncNo = document.getElementById('tbFuncNo').value;
            var objParams = document.getElementById('selectRowPara').value;
            if(CheckChangeParams(objFuncNo,objParams)>0)
            {
                objArrayIDs = GetChkItems();
                if(objArrayIDs.length>0)
                {
                    for(i=0;i<objArrayIDs.length;i++)
	                {
		                if(i==objArrayIDs.length -1){selectIDs +=objArrayIDs[i];}
		                else{selectIDs +=objArrayIDs[i] + "s" ;}
	                }
                    window.location.href = objUrl + "&k="+selectIDs+"&oNa="+objName;
                }
                else
                {
                    alert("操作失败：请首先从列表中点击选择您想要操作的项目！");
                }
            }
        }
    }
}



// 
function goPages(papams,menus,pages,keys,objContainer)
{
    //var k = document.getElementById("searchKey").value;
    //if (keys!=null && keys!="") keys = encodeURIComponent(keys)
    if(papams.length != 0 && pages.length != 0 && menus.length !=0)
    {
        var s = "r="+Math.random()+"&identityKeys="+papams+"&p="+pages+"&k="+keys+"&m="+menus;
        ajaxLoadPage('../getData.asp',s,"Get",objContainer,"")
    }
    else
    { 
        objContainer.value ="Loadding Failed...";
    }
}
//
// ajax调用数据函数; url: 数据处理文件地址,request传入的值，method:get|post,container输出值的对象，funcno函数编号
function ajaxLoadPage(url,request,method,container,funcNo)
{
    var loading_msg='<img src="/skin/images/ajaxing2.gif">'
	container.innerHTML=loading_msg;
    if (!window.XMLHttpRequest) {window.XMLHttpRequest=function (){return new ActiveXObject("Microsoft.XMLHTTP");}     }
	method=method.toUpperCase();
	
	var loader=new XMLHttpRequest;
	
	if (method=='GET')
	{
		urls=url.split("?");
		if (urls[1]=='' || typeof urls[1]=='undefined'){url=urls[0]+"?"+request;}
		else{url=urls[0]+"?"+urls[1]+request;}
		request=null;
	}
    //loader.setHeader("charset","gb2312");
	//alert(url);
	loader.open(method,url,true);
	//loader.setHeader("charset","gb2312");
	

	if (method=="POST")
	{
		loader.setRequestHeader("Content-Type","application/x-www-form-urlencoded;charset=gb2312");
	}

	loader.onreadystatechange=function()
	{
		if (loader.readyState==1){ $(container).innerHTML=loading_msg + " ...";}
		if (loader.readyState==4)
		{
		    if(loader.status==200)
		    {
		        setInnerData(funcNo,loader.responseText,container);
		    }else{ $(container).innerHTML=reTxt;loader = null;}
		}
	}
	loader.send(request);
}
// 获取数据



// 数据展示；funcno 函数编号，returntext 返回值，objcontainer输出文本的对象
function setInnerData(funcNo,returnText,objContainer)
{
    if(returnText.length>0)
    {
        switch (funcNo)
        {
            case "1.1.1":    //检测用户名是否存在
                //if(returnText =="sysManageMain.aspx"){openWinows(returnText);}
               var response = returnText;
 
      if (response=="1") {
       $(objContainer).innerHTML ="<img src='/skin/images/check_error.gif'>&nbsp;用户名已经存在";
      document.getElementById('Submit').disabled = true;
       } 
      else  {
      $(objContainer).innerHTML = "<img src='/skin/images/check_right.gif'>&nbsp;用户名可以使用";
     document.getElementById('Submit').disabled = false;
           }
		 
	            break;
            case "1.1.2":    //检测关键字是否存在
                //if(returnText =="sysManageMain.aspx"){openWinows(returnText);}
               var response = returnText;
 
      if (response=="1") {
       $(objContainer).innerHTML ="<img src='/skin/images/check_error.gif'>&nbsp;关键字已经存在";
      document.getElementById('Submit').disabled = true;
       } 
      else  {
      $(objContainer).innerHTML = "<img src='/skin/images/check_right.gif'>&nbsp;关键字可以使用";
     document.getElementById('Submit').disabled = false;
           }
		 
	            break;
					      
            case "1.1.3":    //检测报社名是否存在
                //if(returnText =="sysManageMain.aspx"){openWinows(returnText);}
               var response = returnText;
 
      if (response=="1") {
       $(objContainer).innerHTML ="<img src='/skin/images/check_error.gif'>&nbsp;报社名已经存在";
      document.getElementById('Submit').disabled = true;
       } 
      else  {
      $(objContainer).innerHTML = "<img src='/skin/images/check_right.gif'>&nbsp;报社名可以使用";
     document.getElementById('Submit').disabled = false;
           }
		 
	            break;
            default:
	            $(objContainer).innerHTML="加载数据失败 ...";
        }
    }
    else
    {
          $(objContainer).innerHTML="<img src='/skin/images/check_error.gif'>&nbsp;不能为空";
		  document.getElementById('Submit').disabled = true;
    }
}



// 资讯列表
function getDataList(papams,menus,pages,keys,orders,objContainer)
{
    if(papams.length != 0 && pages.length != 0 && menus.length !=0)
    {
        var s = "r="+Math.random()+"&identityKeys="+papams+"&p="+pages+"@"+orders+"@"+menus+"&k="+keys;
        ajaxLoadPage('getData.asp',s,"Get",objContainer,"")
    }
    else
    { 
        $(objContainer).innerHTML="加载数据失败 ...";
    }
}

//Email;
function isEmail(s){
	s = trim(s); 
 	var p = /^[_\.0-9a-z-]+@([0-9a-z][0-9a-z-]+\.){1,4}[a-z]{2,3}$/i; 
 	return p.test(s);
}
//Error Msg;
function ErrMsg(s){
    //var s_info = "<table width='100%' border='0' cellspacing='0' cellpadding='0' ><tr><td width='20' height='30' >&nbsp;</td><td width='800' bgcolor='#FFFFCC' style='border-bottom: 1px solid #ff9966;border-top: 1px solid #ff9966;'><font color='#CC3300'>"+s+"</font></td><td  >&nbsp;</td></tr></table>";
    var s_info = "<table width='100%' border='0' cellspacing='0' cellpadding='0' ><tr><td  bgcolor='#FFFFCC' style='border-bottom: 1px solid #ff9966;border-top: 1px solid #ff9966;'><font color='#CC3300'>"+s+"</font></td><td  >&nbsp;</td></tr></table>";
 	return s_info;
}

//根据ID得到对象
function $(ID)
{
    return document.getElementById(ID);
}

function EncodingCN(x)
{
    var ReturnMsg = "";
    if(rsCode(x,1)==239 && rsCode(x,2)==187 && rsCode(x,3)==191)
    {
        ReturnMsg = x;
    }
    else
    {
        var y=rsLength(x),z,i=1,t="";
        while(i<=y)
        {
            z=rsCode(x,i++);
            if(z<128)
            {
                t+=z;
            }
            else
            {
                t+=rsChar(z*256+rsCode(x,i++));
            }
        }
        ReturnMsg = t;
    }
    return ReturnMsg;
}


// 查看详细
function view()
{
    var k = document.getElementById("ClickID").value;
    var m = document.getElementById("K").value;
    if(k.length != 0)
    {
        var url = "base-view.aspx";
        //window.location.href=url +"?k="+ k;
        openWinows(url +"?t=Ysl_&m="+m+"&k="+ k);
    }
    else
    { 
        alert("请单击选择您想要查看的标题行，然后再点击“查看”按钮；或者直接双击标题行也可查看。");
    }
}

// 双击查看
function viewOnDblClick(id)
{
    var m = document.getElementById("K").value;
    var url="base-view.aspx";
    openWinows(url +"?t=Ysl_&m="+m+"&k="+ id);
}

// 新增
function add()
{
    var url = "base-edit.aspx";
    var m = document.getElementById("K").value;
    openWinows(url +"?m="+m+"&act=add&k=0&t=Ysl_");
}
// 编辑
function edit()
{
    var k = document.getElementById("ClickID").value;
    if(k.length != 0)
    {
        var url = "base-edit.aspx";
        var m = document.getElementById("K").value;
        openWinows(url +"?m="+m+"&t=Ysl_&act=edit&k="+ k);
        
    }
    else
    { 
        alert("请单击选择您想要编辑的数据行，然后再点击“编辑”按钮。");
    }
}

// 删除
function del()
{
    //document.form1.v_ItemChk.value =GetChkItems()
    var k = document.getElementById("ClickID").value;
    var m = document.getElementById("K").value;
    if(k.length != 0)
    {
        if(confirm("在删除之前，请您再确认一次，您真的要删除选中的数据吗？")==true)
        {
            ExecSCommand(k,m,"",document.getElementById("MsgInfo"));
            window.location.reload(true);
        }
    }
    else
    { 
        alert("请单击选择您想要删除的数据行，然后再点击“删除”按钮。");
    }
}
