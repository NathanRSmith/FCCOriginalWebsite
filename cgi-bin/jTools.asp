
<SCRIPT LANGUAGE="JAVASCRIPT">
//Start JavaScript Source

ns4	= (document.layers)? true:false
ie4	= (document.all)? true:false
function finalaction(where,place,act,max){		
if (act!=""){									
	if (act=="Remove Image"){					
		var x = document.forms[0]
		var s = x.imagemgtup
		var v = document.all.tags("INPUT")["ximagemgtup"]
		var sx
		var iix = '<%=timg%>'
		if (iix != '-1'){
		if (confirm('Remove Image(s)?')){
		if (iix > '0' ){
		for (ii=0; ii<=iix; ii++){				
				if (s[ii].value!=''){
				s[ii].value=v[ii].value+'<%=server.htmlencode(",")%>'+s[ii].value
				v[ii].disabled=true
				}else{
				s[ii].disabled=true
				v[ii].disabled=true
			}																	
		}
		}else{
		s.value=v.value+'<%=server.htmlencode(",")%>'+s.value
		}
		}else{
			return false;
		}																			
		var strwhere = where
		x.action='imagemgtaction.asp'
		x.method='get'
		x.hideme.value='<%=strpage%>'
		return true;			
	}					
	}																							

if (act=="Add Image"){																		
	var d
	var a
	var b
	var h
	var s
	var icon
	d=document.all.tags("INPUT")["imagemgtup"]
	a=document.all.tags("INPUT")["stralt"]
	s=document.all.tags("INPUT")["Submit1"]
	b=document.all.tags("INPUT")["boothumb"]
	icrop=document.all.tags("SELECT")["icrop"]
	h=document.all.tags("INPUT")["hideme"]
	document.forms[0].method='POST'
	document.forms[0].hideme.value='<%=strpage%>'
	document.forms[0].action="imagemgtaction.asp?icrop="+icrop.value+"&stralt="+a.value+"&boothumb="+b.value+"&Submit1="+s.value+"&hideme="+h.value+"&imagemgtup="+d.value
	document.all.tags("INPUT")["boothumb"].disabled = false
		if (d.value!=''){																		
			if (confirm('UPLOAD IMAGE: ['+d.value+']?')){								
				
			}else{
				document.all.tags("INPUT")["boothumb"].disabled = true
				return false;
			}			
		}else{
			document.all.tags("INPUT")["boothumb"].disabled = true								
			alert("Please Select an Image to Upload")
			return false;	
		}																						
	}																							
	if (act=="Modify Placement"){															
			var x = document.forms[0]
			var strwhere=where
			var f
			var g
			var tt = <%=timg+1%>
			x.action=strwhere
			x.method='get'
			x.hideme.value='<%=strpage%>'
			if (confirm('Modify Image Order?')){
				for (ii=1; ii<=max; ii++){	
						f=document.all.tags("INPUT")["imagemgtup"+ii]
						if (ii<=tt){
						f.disabled=true
						f.name="nolongerneeded"
						f.ID="nolongerneeded"
						f=document.all.tags("INPUT")["ximagemgtup"+ii]
						f.name="imagemgtup"
						f.ID="imagemgtup"
						}
				}
				return true;
				}else{
			return false;
			}
	}																	
}else{														
alert("There were errors. Please submit again.["+act+"]")
}															
}															

function intimgorder(num,nme,max,rem,ID,rid,pre){
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
limg="imagemgtup"+ii
xlimg="x"+limg
a=document.all.tags("INPUT")[limg]
B=document.all.tags("INPUT")[xlimg]
	if (pre!=false){
		if (ii==pre){
					vnum = num
					vimg="image"+vnum
					v="image"+vnum
					d=document.all.tags("IMG")[vimg]
					rimg="rimage"+(pre)
			
					r=document.all.tags("IMG")[rimg]
					r.src=d.src
					r.name=vimg
					r.height=<%=pginfo(3)*.75%>
					r.width=<%=pginfo(2)*.75%>
					r.border=2
			
					d.name=vimg
					d.height=<%=pginfo(3)*.75%>
					d.width=<%=pginfo(2)*.75%>
						var s
						s = getimagefolder(d.src)
						d.src=s+ii+'_TH.gif'
						d.onMouseUp=d.onclick
						d.onclick='#'
					a.value=ii+")"+ nme
					a.ID=rimg
					B.value=rid
				return true;
				}
			}			
	if (rem!=''){
			rimg="rimage"+num
			r=document.all.tags("IMG")[rimg]
			if (a.ID==rimg){
					vimg=r.name
				
					d=document.all.tags("IMG")[vimg]
					d.height=<%=pginfo(3)*.75%>
					d.width=<%=pginfo(2)*.75%>
					d.onclick = d.onMouseUp
					d.onMouseUp = ''
					d.src=r.src
					
					r.src=''
					r.name=''
					r.height=0
					r.border=0
				
					a.value=''
					a.ID=''
					B.value=''
				break
				}
			}
		else{			
		if (a.value=='')
			{
			if (pre==false){
				vnum = num
				vimg="image"+vnum
				v="image"+vnum
				d=document.all.tags("IMG")[vimg]
				rimg="rimage"+(ii)
		
				r=document.all.tags("IMG")[rimg]
				r.src=d.src
				r.name=vimg
				r.height=<%=pginfo(3)*.75%>
				r.width=<%=pginfo(2)*.75%>
				r.border=2
			
				d.name=vimg
				d.height=<%=pginfo(3)*.75%>
				d.width=<%=pginfo(2)*.75%>
					var s
					s = getimagefolder(d.src)
				d.src=s+ii+'_TH.gif'
				d.onMouseUp=d.onclick
				d.onclick='#'
				a.value=ii+")"+ nme
				a.ID=rimg
			
				B.value=rid
			break
			}
			}
		}
	}
}
function getimagefolder(where){
	s = new String(where)
	var z = s.lastIndexOf("/")
	s = s.substr(0,z+1)
	return s;	
}
function highlight(vimg){
d=document.all.tags("IMG")[vimg]
if (d.border=="4"){
d.border="0"
d.vspace="4"
d.hspace="4"
return true
		  }
	else {
d.border="4"
d.vspace="0"
d.hspace="0"
return true
		 }
}

function actionform(){
var	x =	document.forms[0].straction
var	y =	document.forms[0].strpage
if (confirm("You have selected: '"+x.value+"' in '"+y.value+"'.")){
return true
}
	else {
alert("Canceled")
return false
	}
}

function intimgselect(num,total){
var d
var vimg
var strim
var tmpval
var limg
var t
var v
strim = "imagemgrup"
vimg="image"+num
d=document.all.tags("IMG")[vimg]
if (d.style.border=="maroon 2px solid"){
		var go = "new"
	}else{
		go = "old"
}
for(ii=0;ii<=total;ii++){
limg = "imagemgtup"+ii
himg = "ximagemgtup"+ii
t=document.all.tags("INPUT")[limg]
h=document.all.tags("INPUT")[himg]
	if (go=="new"){
		if (t.value==''){
			t.value=d.src
			h.value=d.name
			d.style.border="white 2px solid"
			break
		}	
} else {
	if (t.value==d.src){
		t.value=''
		h.value=''
		d.style.border="maroon 2px solid"
			for (z=ii; z<=total-1; z++){
				limg = "imagemgtup"+z
				himg = "ximagemgtup"+z
				t=document.all.tags("INPUT")[limg]
				limgx = "imagemgtup"+(z+1)
				himgx = "ximagemgtup"+(z+1)
				v=document.all.tags("INPUT")[limgx]
				hv=document.all.tags("INPUT")[himgx]
				t.value=v.value		
				h.value=hv.value					}
				limg = "imagemgtup"+total
				himg = "ximagemgtup"+total
				t=document.all.tags("INPUT")[limg]
				t.value=''
				h.value=''
			break
		}
	}
	}
}

function forcethumbs(act){
	if (act=="1"){
		document.all.tags("INPUT")["boothumb"].checked  = true
		document.all.tags("INPUT")["boothumb"].disabled = true
	}
	else
	{
		document.all.tags("INPUT")["boothumb"].checked  = false
		document.all.tags("INPUT")["boothumb"].disabled = false
	}
}

//END JAVASCRIPT SOURCE
</SCRIPT>
