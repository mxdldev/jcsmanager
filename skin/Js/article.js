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
	if (document.myform.Action.value=="Del")
	{
	   if(confirm("确定要删除选中的文章吗？"))
		 return true;
	   else
		 return false;
	}	 
}
  function CheckInput()
  {
    var Title
	 Title=document.SearchForm.Title.value;
	// alert(PaperName);
	 if(Title=='')
	 {
	   alert('请您输入你所要搜索文章的关键字！');
	   return false;
	 }
	 else
	 {
	   return true;
	 }
  }