<!--#include virtual="includes/nFORMSincludes_contactus.asp"-->
<div class="formLeft">
<div class="formBox">
<form name="formValidation" method="post" action="frame_nFORMAction_contactus.asp">
<% dim validate:validate=gValidate("contactfirstname,contactlastname,contactcity,contactstate,contactzip,contactcountry,contactphone,contactemail,contactperson,comments","First Name,Last Name,City,State,Zip Code,Country,Phone Number,Email Address,Person to Contact,Comments") %>
<% if validate = "" then %>
<% w formText("Step Two: Form Verification","","divFormTitle") %>
<% w gFormInputValidation("html") %>
<% w gImageValidation(constNumberOfImagesForValidation) %>
<% w formText(" ","","divFormText") %>  
<% w gInputOnclick("button","Back to Step One","input inputSide floatLeft","Back to Step One","8","document.hiddenForm.submit();") %>
<% w gInput("submit","Verify Code and Submit","input inputSide floatRight","Verify Code and Submit","8") %>
<% w gHiddenFields(vlist,varray,"noskip")%>
<% else %>
<% w validate %>
<% w formText("","","divFormText") %>
<% w gInputOnclick("button","Back to Step One","input inputSide floatLeft","Back to Step One","8","document.hiddenForm.submit();") %>
<% end if %>
</form>
<form name="hiddenForm" method="post" action="frame_nFORM_contactus.asp">
<% w gHiddenFields(vlist,varray,"skip")%>
</form>
</div>
</div>