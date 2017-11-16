
//此处用来定义控制按钮 请根据自己页面的情况修改等号右边的内容为您相应按钮的id名称
//左边不要修改
var bt=new Array()

bt["cap"]="cap"; //拍摄按钮id
bt["about"]="about";//关于按钮id
bt["again"]="again";//重拍按钮id
bt["save"]="save";//上传保存按钮id

/*
	大头贴控制类 1.0
	您可以通过修改相应的类方法来实现对功能的控制扩展
	*/


	
	
	
	var start=false;
	
	//类块开始 BEGAIN
	
	 function  bigHead(Id){
	
		if (navigator.appName.indexOf("Microsoft") > -1) 
		{ 
				this.swfObj=window[Id]; 
			} else { 
				this.swfObj=document[Id]; 
		}
		
		this.Version="1.00";
		this.PoweredBy="TFOT";
		
	}
	
	//以下内容一般请不要修改(除非您了解其工作过程或对javascript比较熟悉)
	bigHead.prototype.changePic=function (url){if(start==true){this.swfObj.changePic(url)}}
	bigHead.prototype.initialBigHead=function(){this.swfObj.initialBigHead()}
	bigHead.prototype.capturePic=function(){if(start==true){this.swfObj.capturePic()}}
	bigHead.prototype.uploadPic=function(url) {if(start==true){grayAll();this.swfObj.uploadPic(url)}}
	bigHead.prototype.showVersion=function(){
		
			this.swfObj.showVersion();
	}
	
	
	//类块结束 END
	
	function getObj(id){return document.getElementById(id)}
	
	//普通函数开始(此处的函数用于Flash的回调,请自己定义相关函数代码块)
	
	
	
	function uploadOk(fileurl){
		if(confirm("图片上传成功是否查看")){window.open(fileurl,"_blank");}
//		if(confirm("图片上传成功是否查看")){location.href=fileurl;}
		dt.initialBigHead()
		
	}
	
	
	function cameraOk(){
 		start=true;
		setNor("cap");
		setGray("save");
		setGray("again");

	}
	
	//将按钮设置为无效
	function grayAll(){
		
		setGray("cap");
		setGray("save");
		setGray("again");
	}
	
	function cameraFail(){
   		start=false;
		setGray("cap");
		setGray("save");
		setGray("again");
		alert("没有检测到可用摄像头");
		
  		
	}
	function uploadFail(Errinfo)
	{
		alert("上传失败！刷新重来！");
 
	}

	function captureOk()
	{
 
		setGray("cap")
	    setNor("save");
		setNor("again");
		alert("拍照成功！点击“保存”按钮将照片保存到您的相册！")
 
	}
	//设置相应id元素为不可用状态
	function setGray(id)
	{
	getObj(bt[id]).style.filter='alpha(opacity=30) gray';
	getObj(bt[id]).disabled=true;
	}
	//设置相应id元素为正常状态
	function setNor(id)
	{
		getObj(bt[id]).disabled=false;
		getObj(bt[id]).style.filter='';
	}
	
	
var dt=new Object() //注意此处的dt 最好不要修改,如要改动请对应的修改此文件中的dt 和 显示页面body onload中的dt



