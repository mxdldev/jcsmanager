	function OpenWin(str){
		var win_width=window.screen.availwidth-10;
		var win_height=window.screen.availheight-25;
		window.open (str,"", "fullscreen=1,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbar=yes,resizable=no,copyhistory=yes,width="+win_width+",height="+win_height+",top=0,left=0");
	}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
//删除提示函数
function ConfirmDel()
{
	if (document.myform.action.value=="Del")
	{
	   if(confirm("确定要删除选中的报纸吗？"))
		 return true;
	   else
		 return false;
	}	 
}

function ConfirmRev()
 { 
   document.myform.action.value="check"
   var k=0;
   var j
   var strid="";
   var obj=document.getElementsByName("ID");
   for(var i=0;i<obj.length;i++) 
     { 
        if(obj[i].checked==true) 
		 {
            k=k+1; 
			if(i==0)
			 {
			  strid=obj[i].value;
			 }
			 else
			 {
			  strid=strid+","+obj[i].value;
			  } 
         }
	 }
	
    if(k>0)
	 { 
	  //alert(strid);
	 //document.myform.action='FirReview.asp';
	 window.location.href='FirReview.asp?action=check&id='+strid;
      document.myform.submit(); 
	 return true;
	 }
	 else
	 {
	  alert("至少选择一期的报纸");
	   return false;
	 }
}

  function CheckInput()
  {
    var PaperName
	 PaperName=document.SearchForm.PaperName.value;
	// alert(PaperName);
	 if(PaperName=='')
	 {
	   alert('请您输入你所要搜索报纸的关键字！');
	   return false;
	 }
	 else
	 {
	   return true;
	 }
  }
