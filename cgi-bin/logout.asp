<!--#include file="../include/conn.asp"-->
<%
	page = request("page")
	if session("sessionid")<>"" then
		set rs = server.createobject("adodb.recordset")
		sql = "SELECT UserLog.exited FROM UserLog WHERE id=" & session("sessionid") & ";"
		rs.open sql,conn,3,3
		rs.fields("exited") = now
		rs.update
		rs.close
	end if
	session.abandon
	if page="" then
		response.redirect "http://www.jfelectricsupply.com"
	else
		response.redirect page
	end if
%>