<%option explicit%>
<!--#include file="../include/listconn.asp"-->
<% 
	if not session("textupdate") then 
		response.redirect("login.asp")
	end if
	dim sql,rs, i, arrlist
	set rs = server.createobject("adodb.recordset")
	
	

%>

<html>
<head>
	<title>J&F Phone Book</title>
<link rel="stylesheet" href="../include/styles.css">
</head>

<body>
<!--#include file="adminmenu.asp" --> 
<p>
    <%
	sql = "SELECT Contacts.* FROM Contacts ORDER BY COMPANYTXT"
	rs.open sql,conn,3,3
	For i = 0 to rs.recordcount-1
%> 
   
<%
	
	response.write "<b>" & rs("COMPANYTXT") & "</b>" & "<br>"
	response.write rs("NAMETXT") & "<br>"
	response.write rs("PHONE") & "<br>"
	response.write rs("FAX") & "<br><br> "
	
	
%>

<%	
	rs.movenext
	Next
%>
</p>
</body>
</html>
