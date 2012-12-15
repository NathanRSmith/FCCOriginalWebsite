<script language="javascript">
function hello(table,index) {
var new_window = window.open("popup.asp?table="+table+"&index="+index,"Info","location=no,toolbar=no,directories=no,status=no,menubar=no,resizable=yes,scrollbars=yes,width=350,height=400,top=160,left=100");
new_window.focus();
}
</script>

<script language="JavaScript" src="menu.js">
/*
Static Top Menu Script
By Constantin Kuznetsov Jr. (GoldenFox@bigfoot.com) 
Featured on Dynamicdrive.com
For full source code and installation instructions to this script, visit Dynamicdrive.com
*/
</script>
<script language="JavaScript" src="menucontext.js"></script>
<script language="JavaScript">
showToolbar();
</script>
<script language="JavaScript">
'function UpdateIt(){
'if (document.all){
'document.all["MainTable"].style.top = document.body.scrollTop;
'setTimeout("UpdateIt()", 200);
'}
'}
'UpdateIt();
</script>
