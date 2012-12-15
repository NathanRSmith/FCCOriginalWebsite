<script language="JavaScript1.2">
	var zindex=100
	function dropit2(e, whichone){
		//if (window.themenu&&themenu.id!=whichone.id)
		//themenu.style.visibility="hidden"
		themenu=eval(whichone)
		if (document.all){
			themenu.style.left=e.clientX
			themenu.style.top=e.clientY
			if (themenu.style.visibility=="hidden"){
				themenu.style.visibility="visible"
				themenu.style.zIndex=zindex++
			}
			else{
				hidemenu(themenu)
			}
		}
	}
	function dropit(e,whichone){
		//if (window.themenu&&themenu.id!=eval(whichone).id)
		//themenu.visibility="hide"
		themenu=eval(whichone)
		if (themenu.style.visibility=="hidden")
		themenu.style.visibility="visible"
		else
		themenu.style.visibility="hidden"
		themenu.zIndex++
		alert(themenu.style.left)
		//themenu.style.left=e.pageX-e.layerX+6
		//themenu.style.top=e.pageY-e.layerY+19-2
	return false
	}

	function hidemenu(whichone){
		themenu.style.visibility="hidden";
	}

	function hidemenu2(){
		themenu.visibility="hide"
	}

	/*if (document.all){
		document.body.onmouseover=hidemenu
	}*/
</script>

