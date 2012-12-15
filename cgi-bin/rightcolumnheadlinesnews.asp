
<table cellpadding="2">
  <%
	
	set colrs = server.createobject("adodb.recordset")
	sql = "SELECT * FROM PageText WHERE (((PageText.PageID)='" & column & "'));"
	colrs.open sql, conn,3,3
	text = colrs("Text")
	dim pageurl
	pageurl = request.servervariables("PATH_INFO")
	pageurl = split(pageurl,"/")
	if text<>"" then
		text_array = split(text,vbCrLf)
		text_count = ubound(text_array)
		for text_counter = 0 to text_count
			if left(text_array(text_counter),2)="//" then
				header = right(text_array(text_counter),len(text_array(text_counter))-2)
				header = server.htmlencode(header)
				%> 
  <tr>
<td align="right">
  <a class="news" href="<%=pageurl(ubound(pageurl))%>#<%=header%>"><%=header%></a><br><br>
</td>
</tr>
   
    <%
				if not colrs("HeaderCenter") then response.write ""
			end if
		next
	end if
	colrs.close
%> 
</table>
