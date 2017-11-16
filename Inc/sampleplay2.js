var mpUndefined= 0;
var mpStopped= 1;
var mpPaused= 2;
var mpPlaying= 3;
var mpScanForword= 4;
var mpScanReverse= 5;
var mpBuffering= 6;
var mpWaiting= 7;
var mpMediaEnded= 8;
var mpTransitioning= 9;
var mpReady = 10;
var mpReconnection= 11;
var PlayerOldSongID="";
var tid;
var strImgPlay = "images/play/bt_play.gif";
var strImgLoad = "images/play/bt_load.gif";
var strImgStop = "images/play/bt_stop.gif";
var strImgNo= "images/play/bt_no.gif";
var index="";     

//if(document.getElementById("counterimg")==null) document.write('<img width="0" height="0" border="0" src="" id="counterimg"/>');

function callUrl(url)
{

	document.getElementById("counterimg").src = url;
}


function playring(songID,code,ringindex,sType,imgIndex)
{

	var strIndex;
    if(arguments.length==2) 
    {
		strIndex="";
		ringindex="";
	}
	else
		strIndex=",'"+ringindex+"'";
		
	var strType;
    if(arguments.length<4) 
    {
		strType="";
	}
	else
		strType=",'"+sType+"'";
		
	var strImgIndex;
    if(arguments.length<5) 
    {
		strImgIndex="";
	}
	else
	{
		strImgIndex=",'"+imgIndex+"'";
		strImgPlay = "images/play/bt_play"+imgIndex+".gif";
		strImgLoad = "images/play/bt_load"+imgIndex+".gif";
		strImgStop = "images/play/bt_stop"+imgIndex+".gif";
	}
	
	var str;
	if(parseInt(songID)<0)
		str="<IMG alt='暂无试听'  id=\"img_"+songID+ringindex + "\"  src='"+ strImgNo + "' width=14 height=14 border=\"0\">";
	else
		str="<IMG alt='点击播放' style='cursor:hand;' onclick=\"player('"+songID+"','"+code+"'"+strIndex+strType+strImgIndex+");\" id=\"img_"+songID+ringindex + "\" src='" + strImgPlay + "' border=\"0\">";
	//var str="<A onfocus=\"this.blur()\" onclick=\"player("+songID+",'"+code+"'"+strIndex+");\" href=\"javascript:void(0);\"><IMG onclick=\"player("+songID+",'"+code+"'"+strIndex+");\" id=\"img_"+songID+ringindex + "\" src=\"images/play/bt_play.gif\" border=\"0\"></a>";
	document.write(str);
	if(document.getElementById("spanPlayRing")==null)
		document.write("<span id='spanPlayRing'></span>");

}

function player(songID,code,sIndex,sType,imgIndex)
{            

		
		if(arguments.length==5)
		{
			strImgPlay = "images/play/bt_play"+imgIndex+".gif";
			strImgLoad = "images/play/bt_load"+imgIndex+".gif";
			strImgStop = "images/play/bt_stop"+imgIndex+".gif";
		}

	   if(arguments.length==2) 
		sIndex="";
		
	   if(arguments.length<4) 
			sType=1;
			
       tid = songID+sIndex ;
       index=sIndex;       //当前点击的歌曲index
       var streamurl;
	   if(songID.indexOf('//')>-1)
		   streamurl=songID;
	   else
		   streamurl= 'playlist.asp?songid='+songID;  

       if (PlayerOldSongID=="")
	   {		
			PlayerOldSongID=songID+index;
			GE('img_'+tid).src= strImgLoad;
			txt='<object id=ClipPlayer classid=CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6 width=0 height=0>';
			txt= txt + '<param name=autoStart value=false>';
			txt= txt + '<param name=uimode value=invisible>';
			txt= txt + '<param name=enableContextMenu value=false>';
			txt= txt + '<param name=url value='+streamurl+'>';
			txt= txt + '</object>';
			txt= txt + '<script for="ClipPlayer" event="PlayStateChange(NewState)" language="javascript">CRPlayerNewState(NewState);</script>';
			document.all.spanPlayRing.innerHTML = txt;
			
	   }else
	   {
	      if(PlayerOldSongID!=songID+index)
	      {		
	         GE('img_'+PlayerOldSongID).src= GE('img_'+PlayerOldSongID).src.replace('bt_load','bt_play').replace('bt_stop','bt_play');
	         GE('img_'+PlayerOldSongID).alt="点击播放";
	          GE('img_'+songID+index).src= strImgLoad;	       
		     GE('ClipPlayer').url=streamurl;		 				
		     PlayerOldSongID=songID+index;	         
	      }	      
	   }	   

	   if(GE('ClipPlayer').PlayState==mpPlaying || GE('ClipPlayer').PlayState== mpBuffering)
	   {
	     GE('ClipPlayer').controls.stop();
	      GE('ClipPlayer').url="";
	   }
	   else
	   {
	     setTimeout("GE('ClipPlayer').controls.play()",100);
	     //callUrl('http://music.a8.com/counter.aspx?type=0&cntid=' + songID);
	   }


		
}



function CRPlayerNewState(NewState)
{	
	try
	{		
		var strimgUrl=strImgPlay;
		var stralt="";
		
		switch(NewState)
		{
			case mpPlaying:		
				strimgUrl=strImgStop;	
				stralt="点击停止";	
			break;
			case mpBuffering:		
				//return;
				strimgUrl=strImgLoad;				
				stralt="缓冲中...";	
			break;
			case mpStopped:
				strimgUrl=strImgPlay;	
				stralt="点击播放";	
			break;
			case mpReconnection:
				strimgUrl=strImgLoad;	
			break;
			case mpMediaEnded:					
				strimgUrl=strImgPlay;			
			break;
			case mpTransitioning:
				return;
			case mpReady:
				//alert('ok');
				if(PlayerOldSongID!=tid) return;
				stralt="点击播放";	
			break;
		}		
		GE('img_'+tid).src = strimgUrl;
		GE('img_'+tid).alt=stralt;
		
	}
	catch(e)
	{
	}
}
var CurrentPlayImage;

function clpov(item)
{
   
    if(item == CurrentPlayImage)
    {
        if(CurrentPlayImage.children[0].isLoading != true)
            CurrentPlayImage.children[0].src = "play/bt_stop_o.gif";
    }
    else if(item.children[0]) 
        item.children[0].src =  "play/bt_play_o.gif";
}

function clpot(item)
{
   
    if(item == CurrentPlayImage)
    {
        if(CurrentPlayImage.children[0].isLoading != true)
            CurrentPlayImage.children[0].src = "play/bt_stop.gif";
    }
    else if(item.children[0])
        item.children[0].src ="play/bt_play.gif";
       // alert(item.children[0].isLoading);
}


function SF(func) 
{ try 
    { 
       if(eval('window.' + func) != null) 
       { 
          return true; 
       }else
       { 
         return false; 
       } 
   } catch(e) 
   { 
      alert('Failed exec:'+code+'\r\n'+e.description); 
    } return false; 
}
function GE(v) 
{

  if(document.all)
  { 
     return document.all[v]; 
  } 
     return document.getElementById(v); 
}
 