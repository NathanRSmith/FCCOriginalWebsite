<%option explicit%>
<!--#include file="../include/listconn.asp"-->
<% 
	if not session("textupdate") then 
		response.redirect("login.asp")
	end if
	dim sql,rs, i, arrlist
	set rs = server.createobject("adodb.recordset")
	
	if request.form("send")<>"" then
	
	'response.write "here"
	
		arrlist = request.form("address")
		if instr(arrlist,",") then
		arrlist = split(arrlist,",")
		else
		arrlist = array(arrlist)
		end if
	'response.write arrlist(0)
	
	FOR i = 0 to ubound(arrlist)
	sql = "SELECT * FROM Contacts Where ContactID=" & arrlist(i) & ";"
	rs.open sql,conn,3,3
		response.write arrlist(i) & " " & rs("EMAILTXT") & "<br>"
	dim objJmail
	set objJmail = server.createobject("Jmail.Message")

	objJmail.From = "info@jfelectricsupply.com"
	objJmail.FromName = "J&F WebOffers"

	objJmail.AddRecipient rs("EMAILTXT")

	objJmail.Subject = "J&F Web Offers " & FormatDateTime(Date(),2)

	objJmail.Body = request.form("body") &vbCrLF

	objJmail.AppendText("" & vbCrLf)

	
	if not objJmail.Send("mail.visamail.net") then response.write objJmail.log
	objJmail.Clear
	
	'objJmail.Send("mail.visamail.net")
	
	rs.close
	NEXT
	end if

%>

<html>
<head>
	<title>J&F Web Offers and Newsletter</title>
<link rel="stylesheet" href="../include/styles.css">
</head>

<body>
<!--#include file="adminmenu.asp" --> <br>
<font face="Arial, Helvetica, sans-serif" size="2" color="#990000">Select WebOffer 
recipients from the following list. </font> 
<form name="form" action="newsletteradmin.asp" method="post">
  <select name="address" size="6" multiple class="formStyle">
    <%
	sql = "SELECT Contacts.* FROM Contacts ORDER BY EMAILTXT"
	rs.open sql,conn,3,3
	For i = 0 to rs.recordcount-1
%> 
    <option value="<%=rs("contactid")%>">
<%
	'response.write "<font size=""2"" face=""arial"">"
	response.write rs("EMAILTXT")
	'response.write "</font>"
%>
</option>
<%	
	rs.movenext
	Next
%>
</select>
<br>
  <font face="Arial, Helvetica, sans-serif" size="2" color="#990000">WebOffer 
  Details </font> 
  <table>
<tr>
<td colspan="2">
        <textarea name="body" cols="67" rows="15"></textarea>
      </td>
</tr>
<tr>
      <td colspan="2" align="center"> 
        <div align="right">
          <input type="submit" value="Send" name="Send">
        </div>
      </td>
</tr>

</table></form>
</body>
</html>
