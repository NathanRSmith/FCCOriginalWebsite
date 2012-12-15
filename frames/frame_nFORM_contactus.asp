<!--#include virtual="includes/nFORMSincludes_contactus.asp"-->
<div class="formLeft">
<div class="formBox">
<form name="leaseForm" action="frame_nFORMConfirm_contactus.asp" method=post>

<% w formText("Step One: Contact Information","","divFormTitle") %>
<% w formText("First Name","","divFormTextSide floatLeft") %>
<% w formText("&nbsp;&nbsp;Last Name","","divFormTextSide floatLeft") %>

<% w gInput("text","contactfirstname","input inputSide FloatLeft",contactfirstname,"50") %>
<% w gInput("text","contactlastname","input inputSide FloatRight",contactlastname,"50") %>
<% w formTextNewLine() %>

<% w formText("Business Name","(optional)","divFormTextSide floatLeft") %>
<% w formText("&nbsp;&nbsp;Position","(optional)","divFormTextSide floatLeft") %>
<% w formTextNewLine() %>

<% w gInput("text","contactcompany","input inputSide floatLeft",contactcompany,"50") %>
<% w gInput("text","contactposition","input inputSide floatRight",contactposition,"50") %>
<% w formTextNewLine() %>

<% w formText("City","","divFormTextSide floatLeft") %>
<% w formText("&nbsp;&nbsp;State/Province","","divFormTextSide floatRight") %>
<% w formTextNewLine() %>

<% w gInput("text","contactcity","input inputSide floatLeft",contactcity,"50") %>
<% w gSelectState("contactstate","input inputSide floatRight",contactstate) %>
<% w formTextNewLine() %>

<% w formText("Zip Code","","divFormTextSide floatLeft") %>
<% w formText("&nbsp;&nbsp;Country","","divFormTextSide floatRight") %>
<% w formTextNewLine() %>

<% w gInput("text","contactzip","input inputSide floatLeft",contactzip,"10") %>
<% w gSelectCountry("contactcountry","input inputSide floatRight",contactcountry) %>
<% w formTextNewLine() %>

<% w formText("Phone Number","ex. 123-456-7890 xxx","divFormTextSide FloatLeft") %>
<% w formText("&nbsp;&nbsp;Fax Number","(optional)","divFormTextSide FloatRight") %>
<% w formTextNewLine() %>

<% w gInput("text","contactphone","input inputSide floatLeft",contactphone,"20") %>
<% w gInput("text","contactfax","input inputSide floatRight",contactfax,"20") %>
<% w formTextNewLine() %>

<% w formText("Email Address","","divFormText") %>
<% w gInput("text","contactemail","input inputLong",contactemail,"100") %>
<% w formTextNewLine() %>

<% w formText("Person to Contact","","divFormText") %>
<% w gSelectBasic("contactperson","input inputLong",Join(contactarray,","),Join(contactvaluearray,","),contactperson) %>
<% w formTextNewLine() %>

<% w formText("Question/Comment","","divFormText") %>
<% w gTextArea("comments","input inputLong tall100",comments,"255") %>
<% w formTextNewLine() %>

<% w formText("","","divFormText") %>
<% w gInput("submit","Continue to Step Two","input inputSide floatRight","Continue to Step Two","8") %>

<% w gHiddenFields(vlist,varray,"noskip")%>
       
</form>
</div>
</div>