<!--#include file="../include/vHeader.asp"-->
<!--#include file="../include/vTools.asp"-->
<%
'<-- RUN THE USER PERMISSION CHECK
'------------------------------------------------------------------------------------------------------------------------>	
				Call cLogin("textupdate")
'------------------------------------------------------------------------------------------------------------------------>	
'<-- STEP ONE
'------------------------------------------------------------------------------------------------------------------------>	
				Call getPA("strpage","hideme")
'------------------------------------------------------------------------------------------------------------------------>	
'<-- STEP TWO
'------------------------------------------------------------------------------------------------------------------------>	
				Call iHeader()
'------------------------------------------------------------------------------------------------------------------------>	
'<-- STEP THREE
'------------------------------------------------------------------------------------------------------------------------>	
				Call DynamicPage(pgaction)
'------------------------------------------------------------------------------------------------------------------------>	
'<-- STEP FOUR
'------------------------------------------------------------------------------------------------------------------------>	
				Call iFooter(pgselect)
'------------------------------------------------------------------------------------------------------------------------>	
'<-- PRINT HTML
'------------------------------------------------------------------------------------------------------------------------>	
				Response.Write (ihtml)
%>
<!--#include file="../include/vFooter.asp"-->