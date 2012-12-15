<%
	
	set colrs = server.createobject("adodb.recordset")
	sql = "SELECT * FROM PageText WHERE (((PageText.PageID)='" & column & "'));"
	colrs.open sql,conn,3,3
	text = colrs("Text")
	if text<>"" then
		text_array = split(text,vbCrLf)
		text_count = ubound(text_array)
		for text_counter = 0 to text_count
			if left(text_array(text_counter),2)="//" then
				header = right(text_array(text_counter),len(text_array(text_counter))-2)
				header = server.htmlencode(header)
				if colrs("CapsHeader") then header=ucase(header)
				if colrs("HeaderCenter") then response.write "<center>"
				response.write "<font face='" & colrs("HeaderFont") & "' size='" & colrs("HeaderSize") & "' color='" & colrs("HeaderColor") & "'>"
				if colrs("HeaderBold") then response.write "<b>"
				if colrs("HeaderItalics") then response.write "<i>"
				%> 
</p><BR><BR>
<a name="<%=header%>"></a><%=header%><BR>
<!-- 
<HR>	
<TABLE align="right" valign="bottom">
<TR>
	<TD ALIGN="right" valign="bottom"> <h5><a href="#top"><b>BACK 
        TO TOP</b></a></h5></TD>
</TR>
</TABLE><BR><BR>
-->
<p>
<%
				if colrs("HeaderItalics") then response.write "</i>"
				if colrs("HeaderBold") then response.write "</b>"
				response.write "</font>"
				if colrs("HeaderCenter") then response.write "</center>"
				if not colrs("HeaderCenter") then response.write "<br>"
			else
				if colrs("TextCenter") then response.write "<center>"
				response.write "<p>"
				if colrs("TextBold") then response.write "<b>"
				if colrs("TextItalics") then response.write "<i>"
				response.write text_array(text_counter)
				if colrs("TextItalics") then response.write "</i>"
				if colrs("TextBold") then response.write "<b>"
				response.write "</p>"
				if colrs("TextCenter") then response.write "</center>"
				if not colrs("TextCenter") then response.write "<br>"
				
			end if
		next
	end if
	colrs.close
%> 