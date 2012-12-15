<SCRIPT LANGUAGE="JAVASCRIPT">
//Start JavaScript Source
function intimgorder(num,nme,max,rem,ID,rid,pre,alt){
var d
var vimg
var vnum
var rimg
var r
var hx
var h
var l
var xlimg
for(ii=1; ii<=max; ii++){
	if (pre!=false){
		if (ii==pre){
					rimg="rimage"+(pre)
					r=document.all.tags("IMG")[rimg]
					r.src=nme
					r.alt=alt
					r.name=rem
					r.height=<%=pginfo(3)%>
					r.width=<%=pginfo(2)%>
					r.border=2
				return true;
				}
			}			
}
}
function vgpop(isrc,iheight,iwidth,page){
	if (isrc!=''){
		window.setTimeout(window.open('vgdisplay.asp?strpage='+page+'&isrc='+isrc,'ichild','height=100,width=100,titlebar=no,status=no,toolbar=no,alwaysraised=yes,dependent=yes,left=50,top=50,location=no,menubar=no,resizable=no,scrollbars=no,hotkeys=no'),100)
	}
}
function vgshowinfo(what){
	alert(what);
}
//END JAVASCRIPT SOURCE
</SCRIPT>
