<% if ie then %>
<DIV ID="dek" CLASS="dek"></DIV>
<SCRIPT TYPE="text/javascript" language="">
<!--

/*
Pop up information box II (Mike McGrath (mike_mcgrath@lineone.net,  http://website.lineone.net/~mike_mcgrath))
Permission granted to Dynamicdrive.com to include script in archive
For this and 100's more DHTML scripts, visit http://dynamicdrive.com
*/

Xoffset = -90;    // modify these values to ...
Yoffset = 20;    // change the popup position.

var nav,old,iex=(document.all),yyy=-1000;
if(navigator.appName=="Netscape"){(document.layers)?nav=true:old=true;}

if(!old){
var skn=(nav)?document.dek:dek.style;
if(nav)document.captureEvents(Event.MOUSEMOVE);
document.onmousemove=get_mouse;
}

function popup(msg,bak){
var content="<TABLE  WIDTH=175 BORDER=1 BORDERCOLOR=#FF0000 CELLPADDING=1 CELLSPACING=0 "+
"BGCOLOR="+bak+"><TD ALIGN=justify><FONT COLOR=#000000 SIZE=1 face=arial>"+msg+"</FONT></TD></TABLE>";
if(old){alert(msg);return;} 
else{yyy=Yoffset;
 if(nav){skn.document.write(content);skn.document.close();skn.visibility="visible"}
 if(iex){document.all("dek").innerHTML=content;skn.visibility="visible"}
 }
}

function get_mouse(e){
var x=(nav)?e.pageX:event.x+document.body.scrollLeft;skn.left=x+Xoffset;
var y=(nav)?e.pageY:event.y+document.body.scrollTop;skn.top=y+yyy;
}

function kill(){
if(!old){yyy=-1000;skn.visibility="hidden";}
}
//-->
</SCRIPT>
<% end if %>

