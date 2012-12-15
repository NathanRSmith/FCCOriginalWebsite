<%
dim sql, rs, rstemp, conn, param, data, CategoryName, CategoryName2, countrstemp, countrs, counter, counter2, check, Properties, msgBody, index, TableNameCat, TitleOfColumn, TitleOfPage, New_Array, Title, cboBackColor, cboFontColor, txtPageText, ie, browser, cboBodyTextColor
TableNameCat = TableName & "Cat"
%>
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rstemp = Server.CreateObject("ADODB.RecordSet")
'Conn.Open "rvcc"
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Server.MapPath ("/cgi-bin/rvcc.mdb")
%>
<%
	sql = "SELECT * FROM [" & TableName & "] WHERE ((([" & TableName & "].Category)='Information'))"
	rs.open sql, conn, 3, 3
	Properties = rs("Properties")
	rs.close
	Properties=replace(Properties,"|",",,")
	New_Array = split(Properties, ",,")
	TitleOfPage = New_Array(0)
	TitleOfColumn = New_Array(1)
	if ubound(New_Array)>2 then cboBackColor=New_Array(3)
		if cboBackColor="" then cboBackColor="#003366"
	if ubound(New_Array)>3 then cboFontColor=New_Array(4)
		if cboFontColor="" then cboFontColor="#FFFF66"
	if ubound(New_Array)>4 then cboBodyTextColor=New_Array(5)
		if cboBodyTextColor="" then cboBodyTextColor="#FFFF66"
	
%>
<%
	sql = "SELECT Body FROM [" & TableName & "] WHERE ((([" & TableName & "].Category)='Information'))"
	rs.open sql, conn, 3, 3
	txtPageText = rs("Body")
	rs.close
%>
<%
	sql = "SELECT * FROM " & TableName & " WHERE (((" & TableName & ".Category)<>'Information') AND ((" & TableName & ".Locked)='no') AND ((" & TableName & ".TempOff)='no')) ORDER BY " & TableName & ".Category;"
    rs.Open sql, conn, 3, 3
	countrs = rs.recordcount
	countrs = countrs-1		'start count at 0 instead of 1
%>
<%
	sql = "SELECT * FROM " & TableNameCat & " WHERE CategoryName<>'Information'"
    rstemp.Open sql, conn, 3, 3
	countrstemp = rstemp.recordcount
	countrstemp = countrstemp-1		'start count at 0 instead of 1
	rstemp.movefirst
%>
<%
	sub FormatTextarea(text)
		if text<>"" then
			dim text_array, text_count, text_counter, header
			text_array = split(text,vbCrLf)
			text_count = ubound(text_array)
			for text_counter = 0 to text_count
				if left(text_array(text_counter),2)="//" then
					header = right(text_array(text_counter),len(text_array(text_counter))-2)
					response.write "<font face='Arial, Helvetica, sans-serif' size='3' color='" & cboBackColor & "'><b>" & header & "</b></font>"
					response.write "<br>"
				else
					response.write "<font color='" & cboBodyTextColor & "' face='Arial, Helvetica, sans-serif'>" & text_array(text_counter) & "</font><br>"
				end if
			next
		end if			
	end sub
%>
<%
	browser = request.ServerVariables("HTTP_USER_AGENT")
	if instr(browser,"MSIE") then
		ie = true
	else
		ie = false
	end if
%>

