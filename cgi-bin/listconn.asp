<%
	dim conn
	set conn = server.createobject("adodb.connection")
	conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath ("../cgi-bin/mailinglist.mdb")

%>

