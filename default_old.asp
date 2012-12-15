<%
Const Request_POST = 1
Const Request_GET = 2

Set xobj = CreateObject("SOFTWING.ASPtear")
Response.ContentType = "text/html"
    
On Error Resume Next
' URL, action, payload, username, password
strRetval = xobj.Retrieve("http://www.mychurchevents.com/calendar/views/listview.aspx?ci=L6F0I3O9I3I3N8N8H2&list_by=dayspan&DayCount=7&select_by=all_interest_groups&igd=", Request_POST, "test=wille", "", "")

strRetval = Replace(strRetval,"/calendar/","http://www.mychurchevents.com/calendar/")
If Err.Number <> 0 Then
    Response.Write "<b>"
    If Err.Number >= 400 Then
        Response.Write "Server returned error: " & Err.Number
    Else
        Response.Write "Component/WinInet error: " & Err.Description
    End If
    Response.Write "</b>"
    Response.End
End If


Sub w(str)
	response.write str
End sub
%>
<html>
<head>
<title>Welcome to Faith Community Church of Osceola, IN </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--#include file="include/styles.css" -->

</head>
<!-- Start of StatCounter Code -->
<script type="text/javascript">
var sc_project=4022774;
var sc_invisible=1;
var sc_partition=31;
var sc_click_stat=1;
var sc_security="3cc82974";
</script>

<script type="text/javascript"
src="http://www.statcounter.com/counter/counter.js"></script><noscript><div
class="statcounter"><a title="website statistics"
href="http://www.statcounter.com/" target="_blank"><img class="statcounter"
src="http://c.statcounter.com/4022774/0/3cc82974/1/" alt="website
statistics" ></a></div></noscript>
<!-- End of StatCounter Code -->
<body bgcolor="#246C8E" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- ImageReady Slices (homepage2.psd) -->
<table width="781" height="781" border="0" align="center" cellpadding="0" cellspacing="0" id="Table_01">
	<tr>
		<td>
			<img src="images/homepage2_01.gif" width="360" height="33" alt=""></td>
		<td colspan="2">
			<img src="images/homepage2_02.gif" width="420" height="33" alt=""></td>
		<td>
			<img src="images/spacer.gif" width="1" height="33" alt=""></td>
	</tr>
	<tr>
		<td colspan="2" rowspan="9">
			<img src="images/homepage2_03.jpg" width="433" height="457" alt=""></td>
		<td background="images/homepage2_04.jpg"><div align="right"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"><font size="2"> <a href="#preview" class="menu4">Click
        here for a preview of this week's events!&nbsp;&nbsp;</a></font></font> </div></td>
		<td> 
			<img src="images/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_05.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"><a href="welcome.asp" class="menu">WELCOME</a></font></td>
		<td>
			<img src="images/spacer.gif" width="1" height="42" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_06.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"> &nbsp;&nbsp;&nbsp;&nbsp;</font> <a href="whoweare.asp" class="menu">WHO
	    WE ARE</a></td>
		<td>
			<img src="images/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_07.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font> <font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"><a href="whatwedo.asp" class="menu">WHAT
      WE DO                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       </a></font><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
		<td>
			<img src="images/spacer.gif" width="1" height="42" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_08.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font> <font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"><a href="http://www.mychurchevents.com/calendar/calendar.aspx?ci=L6F0I3O9I3I3N8N8H2" target="_blank" class="menu">FULL
      CALENDAR                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       </a></font></td>
		<td>
			<img src="images/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_09.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<a href="contactus.asp" class="menu">CONTACT
      US                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       </a></font></td>
		<td>
			<img src="images/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_10.jpg"><font color="#FF6600" size="5" face="Arial, Helvetica, sans-serif"><a href="global.asp" class="menu">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLOBAL
		      FOCUS</a></font></td>
		<td>
			<img src="images/spacer.gif" width="1" height="42" alt=""></td>
	</tr>
	<tr>
		<td background="images/homepage2_11.jpg"></td>
		<td>
			<img src="images/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	
	<tr>
		<td rowspan="2">
			<img src="images/homepage2_12.jpg" width="347" height="400" alt=""></td>
		<td>
			<img src="images/spacer.gif" width="1" height="126" alt=""></td>
	</tr>
	<tr>
		<td height="274" colspan="2" align="left" valign="bottom" background="images/homepage2_13.jpg" ><div align="center">
          <p><font color="#336699" size="3" face="Arial, Helvetica, sans-serif"><strong>
          Sunday School 9:30 AM<br>
          Coffee &amp; Juice
		          10:25 AM<br>
		      Worship 10:45<!--9:30--> AM<br>
          Wednesdays 6:30 PM<br><br>
		  Check Out Our Blog: <br><a href="http://www.faith-community.blogspot.com/">www.faith-community.blogspot.com</a></strong></font></p>
		  <p>&nbsp;</p>
		</div></td>
		<td>
			<img src="images/spacer.gif" width="1" height="274" alt=""></td>
	</tr>
	<tr>
	  <td height="16" colspan="3" background="/images/calendarcellbg.jpg" style="padding-left:90px; padding-bottom:5px"><div class="textDivCalendar">
<a name="preview"></a>
<table align="center" class="textareaCalendar">
<tr><td bgcolor="#FFFFFF">
<%=strRetval%>
</td>
</tr>
</table>
</div></td>
	  <td>&nbsp;</td>
  </tr>
	<tr>
		<td height="16" colspan="3" background="images/homepage2_14.gif"><div align="center"><font color="#FFCC00" size="2" face="Arial, Helvetica, sans-serif"><strong>Sunday
		      <!--Worship 9:30 &#8226;-->Sunday School 9:30 &#8226; Coffee &amp; Juice
      10:25 &#8226; Wednesdays 6:30 PM </strong></font> </div></td>
		<td>
			<img src="images/spacer.gif" width="1" height="16" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/spacer.gif" width="360" height="1" alt=""></td>
		<td>
			<img src="images/spacer.gif" width="73" height="1" alt=""></td>
		<td>
			<img src="images/spacer.gif" width="347" height="1" alt=""></td>
		<td></td>
	</tr>
</table>
<!-- End ImageReady Slices -->
</body>
</html>