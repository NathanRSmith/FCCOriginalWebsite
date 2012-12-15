<%@Language=VBScript%>
	<% option explicit %>
	<% dim companyname, name, email, address, city, state, zip, phone, fax, conn, rs, continue %>
<%
		set conn = server.createobject("adodb.connection")
		conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath ("cgi-bin/mailinglist.mdb")
		set rs = server.createobject("adodb.recordset")
		rs.open "Contacts",conn,3,3
		rs.addnew
			rs.fields("EMAILTXT") = request.form("email")
			rs.fields("NAMETXT") = request.form("name")
			rs.fields("COMPANYTXT") = request.form("companyname")
			rs.fields("ADDRTXT") = request.form("address")
			rs.fields("CITYTXT") = request.form("city")
			rs.fields("STATETXT") = request.form("state")
			rs.fields("ZIPCODETXT") = request.form("zip")
			rs.fields("PHONE") = request.form("phone")
			rs.fields("FAX") = request.form("fax")
			rs.Fields("OPTIN") = now
			rs.update
		rs.close
		conn.close

		'Declare and set Jmail object
		dim objJmail
		set objJmail = server.createobject("Jmail.Message")
	
		objJmail.From = request.form("email")
		objJmail.FromName = request.form("name")
		objJmail.AddRecipient request.Form("email")
		objJmail.AddRecipientBCC "rnunemaker@jfelectricsupply.com"
		objJmail.AddRecipientBCC "terri@visability.us"
		objJmail.Subject = "WebOffers Registration"
		objJmail.Body = "Company Name : " & request.form("companyname") &vbCrLF
		objJmail.AppendText("Contact Name : " & request.form("name") & vbCrLf)
		objJmail.AppendText("Contact Email : " & request.form("email") & vbCrLf)
		objJmail.AppendText("Phone # : " & request.form("phone") & vbCrLf)
		objJmail.AppendText("Fax # : " & request.form("fax") & vbCrLf)
		objJmail.AppendText("Address : "  & request.form("address") & vbCrLf)
		objJmail.AppendText("City : " & request.form("city") &vbCrLF)
		objJmail.AppendText("State : " & request.form("state") & vbCrLf)
		objJmail.AppendText("Zip Code : " & request.form("zip") & vbCrLf)
		objJmail.AppendText("OptIn Time: " & now)
		'end of objJmail.Body

		objJmail.Send("mail.visamail.net")
	
		'response.redirect("aboutus.asp")
%>
<HTML>
<HEAD>
<TITLE>level2page</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="include/styles.css">



</HEAD><!-- ImageReady Slices (level2page.psd) -->
<BODY bgcolor="#660000" LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>

<TABLE bgcolor="#FFFFFF" background="images/backgroundwhite.gif" WIDTH=793 BORDER=0 CELLPADDING=0 CELLSPACING=0>
  <TR> 
    <TD width="212" ROWSPAN=6 align="left" valign="bottom" height="172"> <IMG SRC="images/level2template_01.jpg" WIDTH=212 HEIGHT=172 ALT=""></TD>
    <TD COLSPAN=4> <IMG SRC="images/level2template_02.jpg" WIDTH=581 HEIGHT=16 ALT=""></TD>
    <TD width="1"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=16 ALT=""></TD>
  </TR>
  <TR> 
    <TD width="3" ROWSPAN=7>&nbsp; </TD>
    <TD>&nbsp; </TD>
    <TD>
      <div align="right"><a href="aboutus.asp"><img src="images/level2template_04.jpg" width=184 height=18 alt="" border="0"></a></div>
    </TD>
    <TD width="42"> <div align="right"><IMG SRC="images/level2template_05.jpg" WIDTH=33 HEIGHT=18 ALT=""></div></TD>
    <TD> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=18 ALT=""></TD>
  </TR>
  <TR> 
    <TD>&nbsp; </TD>
    <TD>
      <div align="right"><a href="linecard.asp"><img src="images/level2template_06.jpg" width=184 height=19 alt="" border="0"></a></div>
    </TD>
    <TD> <div align="right"><IMG SRC="images/level2template_07.jpg" WIDTH=33 HEIGHT=19 ALT=""></div></TD>
    <TD> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=19 ALT=""></TD>
  </TR>
  <TR> 
    <TD>&nbsp; </TD>
    <TD>
      <div align="right"><a href="history.asp"><img src="images/level2template_08.jpg" width=184 height=19 alt="" border="0"></a></div>
    </TD>
    <TD> <div align="right"><IMG SRC="images/level2template_09.jpg" WIDTH=33 HEIGHT=19 ALT=""></div></TD>
    <TD> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=19 ALT=""></TD>
  </TR>
  <TR> 
    <TD width="351">&nbsp; </TD>
    <TD width="185">
      <div align="right"><a href="weboffers.asp"><img src="images/level2template_10.jpg" width=184 height=19 alt="" border="0"></a></div>
    </TD>
    <TD> <div align="right"><IMG SRC="images/level2template_11.jpg" WIDTH=33 HEIGHT=19 ALT=""></div></TD>
    <TD> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=19 ALT=""></TD>
  </TR>
  <TR> 
    <TD COLSPAN=3 ROWSPAN=3 align="left" valign="top">
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr align="left" valign="top"> 
          <td height="216"> 
            <h1>Thank You!</h1>
            <p>You have been registered to receive special offers from J&amp;F 
              Electric. These special web offers are our way of saying thank you 
              to our web users. If you have any questions, feel free to contact 
              us at:</p>
            <blockquote>
              <p>F&amp;F Electric Supply Company, Inc.<br>
                4305 Lincoln Way East<br>
                Mishawaka, IN 46544<br>
                574.256.0599 (phone)<br>
                574.256.9791(fax)<br>
                <a href="mailto:info@jfelectricsupply.com">info@jfelectricsupply.com</a><br>
                www.jfelectricsupply.com </p>
            </blockquote>
            <p>&nbsp;</p>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
      
    </TD>
    <TD height="80"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=81 ALT=""></TD>
  </TR>
  <TR> 
    <TD background="images/level2template_13.jpg" height="350" width="212">&nbsp; </TD>
    <TD height="350"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=360 ALT=""></TD>
  </TR>
  <TR> 
    <TD bgcolor="#660000" align="left" valign="top" height="120"><img src="images/level2template_14.jpg" width=212 height=117 alt=""> 
    </TD>
    <TD height="120"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=117 ALT=""></TD>
  </TR>
</TABLE>
<!-- End ImageReady Slices -->
</BODY>
</HTML>