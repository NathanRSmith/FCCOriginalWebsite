<!--#include file="../include/conn.asp"-->
<% if not session("textupdate") then response.redirect("login.asp") %>
<%
	set rs = server.createobject("adodb.recordset")

	id = request.form("id")

	if request.form("text")<>"" then
		sql = "SELECT PageText.PageID, PageText.Text FROM PageText WHERE (((PageText.PageID)='" & id & "'));"
		rs.open sql,conn,3,3
			rs.fields("PageID") = id
			rs.fields("Text") = request.form("text")
		rs.update
		rs.close
		response.write "<font face='arial' size='4' color='#003366'><b>Changes for the " & id & " page have been made</b></font>"
		id=""
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>J&F Electric Administration</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<META HTTP-EQUIV="expires" CONTENT="0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<form action="file:///H%7C/websites/E-normous/forms/updatetext.asp" method="post">
<% if id="" then %>
<select name="id">
<%
	sql = "SELECT PageText.* FROM PageText INNER JOIN (Personnel INNER JOIN PageRights ON Personnel.PersonnelID = PageRights.PersonnelID) ON PageText.PageID = PageRights.PageID WHERE (((Personnel.Username)='" & session("user") & "')) ORDER BY PageText.Name;"
	rs.open sql,conn,3,3
	do while not rs.eof
	%>
	<option value="<%=rs("PageID")%>"><%=rs("Name")%></option>
	<%
		rs.movenext
	loop
%>
</select>
<input type="submit" value="Edit information">
<% else %>
<%
	sql = "SELECT PageText.* FROM PageText WHERE (((PageText.PageID)='" & id & "'));"
	rs.open sql,conn,3,3
%>
<table cellspacing="0" cellpadding="0" border="0">
	<!-- <tr>
		<td valign="top" align="right"><font face="arial" size="2"><b>
			Header Text Color:&nbsp;&nbsp;
		</td>
		<td valign="top">
			<input type="text" name="headercolor" value="<%=rs("HeaderColor")%>" size="4">
		</td>
	</tr>
	<tr>
		<td valign="top" align="right"><font face="arial" size="2"><b>
			Body Text Color:&nbsp;&nbsp;
		</td>
		<td valign="top">
			<input type="text" name="textcolor" value="<%=rs("TextColor")%>" size="4">
		</td>
	</tr> -->
	<tr>
		
      <td valign="top" colspan="2"><font face="Arial, Helvetica, sans-serif" size="3" color="#000099"><b> 
        Making changes to:</b> <%=rs("Name")%> </font> </td>
	</tr>
	<tr>
		
      <td valign="top" align="right"> <font face="Arial, Helvetica, sans-serif" size="3">Text:&nbsp;&nbsp;</font> 
      </td>
		<td valign="top">
			<textarea name="text" cols="50" rows="20"><%=rs("Text")%></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<p>
			&nbsp;<br>
			<input type="hidden" name="id" value="<%=id%>">
			<center>
			<input type="button" value="Back" onclick="history.back(-1);">
			<input type="reset" value="Reset">
			<input type="submit" value="Submit Information"></center>
			</p>
		</td>
	</tr>
</table>
<% rs.close %>
<% end if %>
</form>
</BODY>
</HTML>
