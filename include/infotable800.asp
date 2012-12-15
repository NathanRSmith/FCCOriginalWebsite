<%
	dim eventon
	sub DateReturn(don,doff)
		dim dateon, dateoff
		eventon=false
		if doff="" or doff="none" then
			dateon = DateDiff("d",date,don)
			if dateon<=0 then eventon=true
		else
			dateon = DateDiff("d",date,don)
			dateoff = DateDiff("d",date,doff)
			if dateon<=0 and dateoff>0 then eventon=true
		end if
	end sub
%>
<%
Sub DrawTable(width)

dim table_width
if width="" then
	table_width=191
else
	table_width=cint(width)
end if	
%>
<table cellspacing="0" cellpadding="9" border="0" width="<%=table_width%>">
	<tr>
		<td bgcolor="<%=cboBackColor%>">
			<center>
			<font face="arial" color="<%=cboFontColor%>">
			<b><%=TitleOfColumn%></b>
			</font>
			</center>
		</td>
	</tr>
	<tr>
		<td>
			<font face="arial">
			<%
				rstemp.movefirst
				for counter = 0 to countrstemp
					CategoryName = rstemp("CategoryName")
					check = 0		'checks to see if category has any entries
					if not rs.eof then
						for counter2 = 0 to countrs
							CategoryName2 = rs("Category")
							call DateReturn(rs("DateOn"),rs("DateOff"))
							if eventon=false then CategoryName2="nothing"
							if CategoryName=CategoryName2 then
								if check = 0 then response.write "<font color='" & cboFontColor & "'><b>" & CategoryName & "</b><br>"
								check = 1		'category has an entry
								Title = replace(rs("Title"),"''","'")
								if len(Title)>32 then Title=left(Title,30) & "..."
								msgBody = rs("Body")
								msgBody=replace(msgBody,"'","")
								msgBody=replace(msgBody,vbCrLf," ")
								if len(msgBody)>150 then msgBody=left(msgBody,150) & "..."
								%>
			<font size=2> <font color="<%=cboBackColor%>">-</font> <a class="table" href="Javascript:hello('<%=TableName%>','<%=rs("index")%>')" <% if ie then %>onmouseover="popup('<%=msgBody%>','#FFFFFF')"; onmouseout="kill()"<% end if %>><%=Title%></a></font><br>
								<%
							end if
							rs.movenext
						next
						rs.movefirst
					end if
					rstemp.movenext
				next
			%>
		</td>
	</tr>
</table>

<%
	rs.close
	rstemp.close
%>
<%
End Sub
%>
<%
Sub DrawTableDropDown(width)

dim table_width
if width="" then
	table_width=191
else
	table_width=cint(width)
end if
%>
<table cellspacing="0" cellpadding="9" border="0" width="<%=table_width%>">
	<tr>
		<td bgcolor="<%=cboBackColor%>">
			<center>
			<font face="arial" color="<%=cboFontColor%>">
			<b><%=TitleOfColumn%></b>
			</font>
			</center>
		</td>
	</tr>
	<tr>
		<td>
			<font face="arial">
			<%
				rstemp.movefirst
				for counter = 0 to countrstemp
					CategoryName = rstemp("CategoryName")
					check = 0		'checks to see if category has any entries
					if not rs.eof then
						for counter2 = 0 to countrs
							CategoryName2 = rs("Category")
							call DateReturn(rs("DateOn"),rs("DateOff"))
							if eventon=false then CategoryName2="nothing"
							if CategoryName=CategoryName2 then
								if check = 0 then response.write "<font color='" & cboFontColor & "'><b>" & CategoryName & "</b></font><br>"
								check = 1		'category has an entry
								Title = replace(rs("Title"),"''","'")
								if len(Title)>32 then Title=left(Title,30) & "..."
								msgBody = rs("Body")
								if len(msgBody)>150 then msgBody=left(msgBody,150) & "..."
								%>
			<font face="Arial" size="2">
			<nobr> <font color="<%=cboBackColor%>">-</font> 
			<a class="table" href="Javascript:hello('<%=TableName%>','<%=rs("index")%>')" onmouseout="hidemenu('dropmenu<%=rs("index")%>')" onMouseOver="dropit2(event, 'dropmenu<%=rs("index")%>')"><%=replace(rs("Title"),"''","'")%></a>
			</nobr>
			</font>
			<br>
			<div id="dropmenu<%=rs("index")%>" style="position:absolute; left:0; top:0; layer-background-color:#FFFFFF; width:170; background-color:#FFFFFF; visibility:hidden; border:1px solid #FF0000; padding:3px;z-index:1;">
			<font color="<%=cboBackColor%>" face="arial" size="1">&nbsp;<br><%=msgBody%></font></div>
								<%
							end if
							rs.movenext
						next
						rs.movefirst
					end if
					rstemp.movenext
				next
			%>
		</td>
	</tr>
</table>
<%
	rs.close
	rstemp.close
%>
<%
End Sub
%>
