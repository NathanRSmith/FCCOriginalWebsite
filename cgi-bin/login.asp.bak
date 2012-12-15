<!--#include file="../include/conn.asp"-->
<%
	if request.form("txtUsername")<>"" and request.form("txtPassword")<>"" then
		txtUsername = request.form("txtUsername")
		txtPassword = request.form("txtPassword")
	%>
	<%
		set rs = server.createobject("adodb.recordset")
		sql = "SELECT Personnel.* FROM Personnel WHERE (((Personnel.Username)='" & txtUsername & "'));"
		rs.open sql,conn,3,3
		if rs.eof then
			logon = false
		else
			if txtPassword <> rs("Password") then
				logon = false
				session("boardlogin")=false
				response.redirect("login.asp")
			else
				logon = true
				session("textupdate")=true
				session("pid")=rs("PersonnelID")
				session("user")=rs("Username")
				response.redirect("newsletteradmin.asp")
			end if
		end if
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>J&F Electric Administration Login</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<META HTTP-EQUIV="expires" CONTENT="0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<p><font face="Arial, Helvetica, sans-serif" size="4" color="003366"><b><u><font color="#990000">J&amp;F 
  Electric Administration Login</font></u><font color="#990000">:</font></b></font></p>
<form name="frmLogin" method="post" action="login.asp">
  <table width="33%" border="0">
    <tr> 
      <td><font face="Arial, Helvetica, sans-serif"><b>Username:</b></font></td>
      <td> 
        <input type="text" name="txtUsername">
      </td>
    </tr>
    <tr> 
      <td><font face="Arial, Helvetica, sans-serif"><b>Password:</b></font></td>
      <td> 
        <input type="password" name="txtPassword">
      </td>
    </tr>
    <tr> 
      <td width="100"> 
        <div align="center"> </div>
      </td>
      <td> 
        <input type="hidden" name="action" value="Login2">
        <input type="submit" name="Submit" value="Submit">
      </td>
    </tr>
  </table>
</form>
</BODY>
</HTML>
