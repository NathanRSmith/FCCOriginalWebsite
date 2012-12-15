<!--#include virtual="includes/nFORMSincludes_contactus.asp"-->
<div class="formLeft">
<div class="formBox">
<form name="formValidation" method="post">
<% dim validate:validate=gImageValidate(constNumberOfImagesForValidation)%>
<% if validate = "" then %>
<% 'submitFormDb() %>
<% w submitFormEmail %>
<% w formText("&nbsp;","","divFormTitle") %>
<% w formText(" ","","divFormText") %>  
<% else %>
<% w validate %>
<% w formText("","","divFormText") %>
<% w gInputOnclick("button","Back to Step Two","input inputSide floatLeft","Back to Step Two","8","document.hiddenForm.submit();") %>
<% end if %>
</form>
<form name="hiddenForm" method="post" action="frame_nFORMConfirm_contactus.asp">
<% w gHiddenFields(vlist,varray,"skip")%>
</form>
</div>
</div>