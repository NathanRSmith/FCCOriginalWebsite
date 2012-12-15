<%
    sub w(str)
        response.Write(str)
    end sub
%>
<%
    function gSelectBasicMultiple(fname,fclass,foptionlist,fvaluelist,flistsize,fselected)
    gSelectBasicMultiple = ""
    gSelectBasicMultiple = "<SELECT MULTIPLE SIZE="""&flistsize&""" name="""&fname&""" id="""&fname&""" class="""&fclass&""">"
                
                if fselected <> "" then
                    fselected = ","&fselected&","
                end if
                if instr(foptionlist,",") then
                    foptionlist = split(foptionlist,",")
                else
                    foptionlist = Array(foptionlist)
                end if
                if instr(fvaluelist,",") then
                    fvaluelist = split(fvaluelist,",")
                else
                    fvaluelist = Array(fvaluelist)
                end if
                For i = 0 to ubound(foptionlist)
                    gSelectBasicMultiple = gSelectBasicMultiple & "<OPTION value="""&fvaluelist(i)&""""

                    fselected = replace(fselected,", ",",")
                    if instr(cstr(fselected),","&cstr(fvaluelist(i))&",") then
                        gSelectBasicMultiple = gSelectBasicMultiple & " selected"
                    end if
                    
                    gSelectBasicMultiple = gSelectBasicMultiple & ">"
                    gSelectBasicMultiple = gSelectBasicMultiple & foptionlist(i)
                    gSelectBasicMultiple = gSelectBasicMultiple & "</OPTION>"
                Next
    gSelectBasicMultiple = gSelectBasicMultiple & "</SELECT>"
    end function

    function gSelectBasic(fname,fclass,foptionlist,fvaluelist,fselected)
    gSelectBasic = ""
    gSelectBasic = "<SELECT name="""&fname&""" id="""&fname&""" class="""&fclass&""">"
                if instr(foptionlist,",") then
                    foptionlist = split(foptionlist,",")
                else
                    foptionlist = Array(foptionlist)
                end if
                if instr(fvaluelist,",") then
                    fvaluelist = split(fvaluelist,",")
                else
                    fvaluelist = Array(fvaluelist)
                end if
                For i = 0 to ubound(foptionlist)
                    gSelectBasic = gSelectBasic & "<OPTION value="""&fvaluelist(i)&""""
                    if cstr(fvaluelist(i)) = cstr(fselected) then
                        gSelectBasic = gSelectBasic & " selected"
                    end if
                    gSelectBasic = gSelectBasic & ">"
                    gSelectBasic = gSelectBasic & foptionlist(i)
                    gSelectBasic = gSelectBasic & "</OPTION>"
                Next
    gSelectBasic = gSelectBasic & "</SELECT>"
    end function
%>
<%
    function gSelectState(fname,fclass,fselected)
    dim astatelist:astatelist="Alabama,Alaska,Arizona,Arkansas,California,Colorado,Connecticut,Delaware,District of Columbia,Florida,Georgia,Hawaii,Idaho,Illinois,Indiana,Iowa,Kansas,Kentucky,Louisiana,Maine,Maryland,Massachusetts,Michigan,Minnesota,Mississippi,Missouri,Montana,Nebraska,Nevada,New Hampshire,New Jersey,New Mexico,New York,North Carolina,North Dakota,Ohio,Oklahoma,Oregon,Pennsylvania,Rhode Island,South Carolina,South Dakota,Tennessee,Texas,Utah,Vermont,Virginia,Washington,West Virginia,Wisconsin,Wyoming,,Alberta,British Columbia,Manitoba,New Brunswick,Newfoundland,Northwest Territories,Nova Scotia,Ontario,Prince Edward Island,Quebec,Saskatchewa,Yukon Territory,,Other"
    astatelist = split(astatelist,",")
    gSelectState = ""
    gSelectState = "<SELECT name="""&fname&""" id="""&fname&""" class="""&fclass&""">"
                For i = 0 to ubound(astatelist)
                    gSelectState = gSelectState & "<OPTION value="""&astatelist(i)&""""
                    if cstr(astatelist(i)) = cstr(fselected) and cstr(fselected)<>"" then
                        gSelectState = gSelectState & " selected"
                    end if
                    gSelectState = gSelectState & ">"
                    gSelectState = gSelectState & astatelist(i)
                    gSelectState = gSelectState & "</OPTION>"
                Next
    gSelectState = gSelectState & "</SELECT>"
    end function
    
    function gSelectCountry(fname,fclass,fselected)
    dim acountrylist:acountrylist="United States,Canada,,Afghanistan,Albania,Algeria,American Samoa,Andorra,Angola,Anguilla,Antarctica,Antigua and Barbuda,Argentina,Armenia,Aruba,Australia,Austria,Azerbaidjan,Bahamas,Bahrain,Bangladesh,Barbados,Belarus,Belgium,Belize,Benin,Bermuda,Bhutan,Bolivia,Bosnia-Herzegovina,Botswana,Bouvet Island,Brazil,Brunei Darussalam,Bulgaria,Burkina Faso,Burundi,Cambodia,Cameroon,Canada,Cape Verde,Cayman Islands,Central African Republic,Chad,Chile,China,Christmas Island,Cocos Islands,Colombia,Comoros,Congo,Cook Islands,Costa Rica,Croatia,Cuba,Cyprus,Czech Republic,Denmark,Djibuti,Dominica,Dominican Republic,East Timor,Ecuador,Egypt,El Salvador,Equatorial Guinea,Estonia,Ethiopia,Falkland Islands,Faroe Islands,Fiji,Finland,France,French Guyana,Gabon,Gambia,Georgia,Germany,Ghana,Gibraltar,Great Britain,Greece,Greenland,Grenada,Guadeloupe,Guam,Guatemala,Guinea,Guinea Bissau,Guyana,Haiti,Honduras,Hong Kong,Hungary,Iceland,India,Indonesia,Iran,Iraq,Ireland,Israel,Italy,Ivory Coast,Jamaica,Japan,Jordan,Kazakhstan,Kenya,Kiribati,Kuwait,Kyrgyzstan,Laos,Latvia,Lebanon,Lesotho,Liberia,Libya,Liechtenstein,Lithuania,Luxembourg,Macau,Macedonia,Madagascar,Malawi,Malaysia,Maldives,Mali,Malta,Marshall Islands,Martinique,Mauritania,Mauritius,Mayotte,Mexico,Micronesia,Moldavia,Monaco,Mongolia,Montserrat,Morocco,Mozambique,Myanmar,Namibia,Nauru,Nepal,Netherlands,Netherlands Antilles,New Caledonia,New Zealand,Nicaragua,Niger,Nigeria,Niue,Norfolk Island,North Korea,Norway,Oman,Pakistan,Palau,Panama,Papua New Guinea,Paraguay,Peru,Philippines,Pitcairn Island,Poland,Polynesia,Portugal,Puerto Rico,Qatar,Reunion,Romania,Russian Federation,Rwanda,Saint Helena,Saint Lucia,St Vincent and Grenadines,Samoa,San Marino,Saudi Arabia,Senegal,Seychelles,Sierra Leone,Singapore,Slovak Republic,Slovenia,Solomon Islands,Somalia,South Africa,South Korea,Spain,SriLanka,Sudan,Suriname,Swaziland,Sweden,Switzerland,Syria,Tadjikistan,Taiwan,Tanzania,Thailand,Togo,Tokelau,Tonga,Trinidad and Tobago,Tunisia,Turkey,Turkmenistan,Tuvalu,Uganda,Ukraine,United Arab Emirates,United Kingdom,United States,Uruguay,Uzbekistan,Vanuatu,Vatican City State,Venezuela,Vietnam,Virgin Islands(British),Virgin Islands(USA),Wallis and Futuna Islands,WesternSahara,Yemen,Yugoslavia,Zaire,Zambia,Zimbabwe"
    acountrylist = split(acountrylist,",")
    gSelectCountry = ""
    gSelectCountry = "<SELECT name="""&fname&""" id="""&fname&""" class="""&fclass&""">"
                For i = 0 to ubound(acountrylist)
                    gSelectCountry = gSelectCountry & "<OPTION value="""&acountrylist(i)&""""
                    if (cstr(acountrylist(i)) = cstr(fselected)) and (cstr(fselected)<>"") then
                    gSelectCountry = gSelectCountry & " selected"
                    end if
                    gSelectCountry = gSelectCountry & ">"
                    gSelectCountry = gSelectCountry & acountrylist(i)
                    gSelectCountry = gSelectCountry & "</OPTION>"
                Next
    gSelectCountry = gSelectCountry & "</SELECT>"
    end function
%>
<%  
    function GenerateRandomValue(fnum)
        dim randomNumber
        GenerateRandomValue = ""
        randomize()
        For grcn = 0 to fnum
        randomNumber = Int(9 * rnd())
        GenerateRandomValue = GenerateRandomValue & randomNumber
        Next
    end function
%>
<%
    function formText(ftext,fhelp,fclass)
        formText = "<div class="""&fclass&""">" & ftext & "<span class=""helpText"">"&fhelp&"</span></div>"
    end function
    
    function formTextNewLine()
        formTextNewLine = "<div class=""divFormNewLine"">&nbsp;</div>"
    end function
%>
<%
    function gSelectDb(fname,fclass,ftable,foption,fvalue,fselected)
    gSelectDb = ""
    gSelectDb = "<SELECT name="""&fname&""" id="""&fname&""" class="""&fclass&""">"
        sql = "SELECT * FROM " & ftable & ";"
        set conn = server.CreateObject("adodb.connection")
            conn.open databaseConnectionString
        set rs = server.CreateObject("adodb.recordset")
            rs.Open sql,conn,3,3
                For i = 1 to rs.RecordCount
                    gSelectDb = gSelectDb & "<OPTION value="""&rs(fvalue)&""""
                    if cstr(rs(fvalue)) = cstr(fselected) then
                        gSelectDb = gSelectDb & " selected"
                    end if
                    gSelectDb = gSelectDb & ">"
                    gSelectDb = gSelectDb & rs(foption)
                    gSelectDb = gSelectDb & "</OPTION>"
                    rs.MoveNext
                Next
            rs.Close
            conn.Close
        set rs = nothing
        set conn = nothing
    gSelectDb = gSelectDb & "</SELECT>"
    end function
%>
<%
    function gRecordDb(ftable,fColumn,fLookup,fvalue)
    gRecordDb = ""
        sql = "SELECT "& fColumn &" FROM " & ftable & " WHERE " & fLookup &" = " & fvalue & ";"
        set conn = server.CreateObject("adodb.connection")
            conn.open databaseConnectionString
        set rs = server.CreateObject("adodb.recordset")
            rs.Open sql,conn,3,3
                if not rs.eof then
                gRecordDb = rs.fields(fColumn)
                end if
            rs.Close
            conn.Close
        set rs = nothing
        set conn = nothing
    end function
%>
<%
    function gInputOnclick(ftype,fname,fclass,fvalue,flength,fscript)
        gInputOnclick = "<input type="""&ftype&""" name="""&fname&""" id="""&fname&""" class="""&fclass&""" value="""&fvalue&""" maxlength="""&flength&""" onclick="""&fscript&""" />"
    end function
%>
<%
    function gInput(ftype,fname,fclass,fvalue,flength)
        gInput = "<input type="""&ftype&""" name="""&fname&""" id="""&fname&""" class="""&fclass&""" value="""&fvalue&""" maxlength="""&flength&""" />"
    end function
%>
<%
    function gTextarea(fname,fclass,fvalue,flength)
        gTextarea = "<textarea name="""&fname&""" id="""&fname&""" class="""&fclass&""" maxlength="""&flength&""">"&fvalue&"</textarea>"
    end function
%>
<%
    function gImage(fname,fclass,fsource)
        gImage = "<img name="""&fname&""" id="""&fname&""" class="""&fclass&""" src="""&fsource&""" />"
    end function
%>
<%
    function gImageValidation(fnum)
        gImageValidation = ""
        dim valimage
        For i = 1 to fnum
        valimage=ImageValidationArray(GenerateRandomValue(0))
        gImageValidation = gImageValidation & gImage("image"&i," ","/images/validation/"&valimage&".gif")
        gImageValidation = gImageValidation & gInput("hidden","h_image"&i,"",valimage,"20")
        Next
        gImageValidation = gImageValidation & formText("Please Enter the Validation Code as Displayed Above","","divFormText")
        gImageValidation = gImageValidation & gInput("text","validationAnswer","input inputLong clearBoth","",fnum)
    end function
%>
<%
    function gImageValidate(fnum)
        dim validkey:validkey=""
        gImageValidate = ""
        for i = 1 to fnum
            for ii = 0 to ubound(ImageValidationArray)
                if request.Form("h_image"&i) = ImageValidationArray(ii) then
                    validkey = validkey & ii
                    exit for
                end if
            next
        next
        if validkey <> request.Form("validationAnswer") then
            gImageValidate = gImageValidate & formText("Error: The Verification Code Did Not Match","","divFormText errorText")
        end if
    end function
%>
<%  
    function gValidate(fvalue,fdisplay)
        gValidate = ""
        if request.Form("h_skip") <> "skip" then
        dim avalue,checkvalue,adisplay
        if instr(fvalue,",") then
            avalue = split(fvalue,",")
        else
            avalue = array(fvalue)
        end if
        if instr(fdisplay,",") then
            adisplay = split(fdisplay,",")
        else
            adisplay = array(fdisplay,",")
        end if
        for i = 0 to ubound(avalue)
            checkvalue = request.Form(avalue(i))
            if len(checkvalue) = 0 then
                gValidate = gValidate & formText("Error: "&adisplay(i),"(Required Field)","divFormText errorText")
            end if
        next
        end if
    end function
    
    function gPasswordMatch(fpassword,fconfirm)
        gPasswordMatch = ""
        if request.Form(fpassword) <> request.Form(fconfirm) then
            gPasswordMatch = gPasswordMatch & formText("Error: Passwords Do Not Match","","divFormText errorText")
        end if
    end function
%>
<%
    function lv(fname)
        dim newname:newname = false
        lv = request.Form(fname)
        if lv = "" then
            For each nnn in request.Form
                if lcase(nnn) = lcase(fname) then
                    newname = false
                exit for
                else
                    newname = true
                end if
            Next
            if newname = true then
                fname = "h_" & fname
            end if
        end if
        lv = request.Form(fname)
    end function
    function bv(fname)
        dim newname:newname = false
        bv = request.QueryString(fname)
        if bv = "" then
            bv = lv(fname)
        end if
    end function
%>
<%
    function gHiddenFields(flist,farray,fskip)
        gHiddenFields = ""
        if fskip = "skip" then
            gHiddenFields = gHiddenFields & gInput("hidden","h_skip","","skip","255")
        end if
        dim alist
        if flist <> "" then
            if instr(flist,",") then
                alist = split(flist,",")
            else
                alist = array(flist)
            end if
        for i = 0 to ubound(alist)
        gHiddenFields = gHiddenFields & gInput("hidden","h_"&alist(i),"",farray(i),"255") 
        next
        end if
    end function
%>
<%
    function gFormInputValidation(ftype)
    gFormInputValidation = ""
    if lcase(ftype) = "html" then
        gFormInputValidation = gFormInputValidation & "<div class=""reviewDiv"">"
        gFormInputValidation = gFormInputValidation & "<table class=""reviewTable"">"
        gFormInputValidation = gFormInputValidation & "<tr><td>"
    end if
        dim fivarray(2)
        commonlist = split(commonlist,",")
        for i = 0 to ubound(varray)
            if instr(lcase(validateskiplist),","&lcase(commonlist(i))&",") then
                if lcase(ftype) = "html" then
                gFormInputValidation = gFormInputValidation & formText(commonlist(i) & ": ","********","divFormText")
                else
                gFormInputValidation = gFormInputValidation & commonlist(i) & ": ********" & vbCrLf
                end if
            elseif instr(lcase(validateskiplist),",[contactperson]"&lcase(commonlist(i))&",") then
                if lcase(ftype) = "html" then
                gFormInputValidation = gFormInputValidation & formText(commonlist(i) & ": ",contactarray(varray(i)),"divFormText")
                else
                gFormInputValidation = gFormInputValidation & commonlist(i) & ": " & contactarray(varray(i)) & vbCrLf
                end if
            else
                if instr(lcase(validatedblookup),","&lcase(commonlist(i))&",") then
                    select case lcase(commonlist(i))
                    case "state"
                        fivarray(0) = "state"
                        fivarray(1) = "StateCode"
                        fivarray(2) = "StateID"
                    case "title"
                        fivarray(0) = "contacttype"
                        fivarray(1) = "contacttypetitle"
                        fivarray(2) = "contacttypeid"
                    case "company type"
                        fivarray(0) = "tenanttype"
                        fivarray(1) = "tenanttypename"
                        fivarray(2) = "tenanttypeid"
                    case "approximate square footage"
                        fivarray(0) = "squarefootage"
                        fivarray(1) = "squarefootagetext"
                        fivarray(2) = "squarefootageid"
                    end select
                    if lcase(ftype) = "html" then
                    gFormInputValidation = gFormInputValidation &  formText(commonlist(i) & ": ",gRecordDb(fivarray(0),fivarray(1),fivarray(2),varray(i)),"divFormText")
                    else
                    gFormInputValidation = gFormInputValidation &  commonlist(i) & ": " & gRecordDb(fivarray(0),fivarray(1),fivarray(2),varray(i)) & vbCrLf
                    end if
                else
                    if lcase(ftype) = "html" then
                    gFormInputValidation = gFormInputValidation &  formText(commonlist(i) & ": ",varray(i),"divFormText")
                    else
                    gFormInputValidation = gFormInputValidation &  commonlist(i) & ": " & varray(i) & vbCrLf
                    end if
                end if
            end if
        next
        if lcase(ftype) = "html" then
        gFormInputValidation = gFormInputValidation & "</td></tr>"
        gFormInputValidation = gFormInputValidation & "</table>"
        gFormInputValidation = gFormInputValidation & "</div>"
        end if
    end function
%>
<%
    function submitformemail()
        set osmtp = server.CreateObject("ansmtp.obj")
        osmtp.ServerAddr = constmailserver
        osmtp.Subject = constcontactformsubjectline
        osmtp.FromAddr = constcontactformfromaddress
        osmtp.From = constcontactformfromaddress

        dim aconstcontactformrecipient:aconstcontactformrecipient = contactemailarray(contactperson)
        if instr(aconstcontactformrecipient,",") then
        aconstcontactformrecipient = split(aconstcontactformrecipient,",")
        else
        aconstcontactformrecipient = Array(aconstcontactformrecipient)
        end if
        For i = 0 to ubound(aconstcontactformrecipient)
        osmtp.AddRecipient aconstcontactformrecipient(i),aconstcontactformrecipient(i),"0"
        Next
        
        osmtp.BodyText = gFormInputValidation("text") & vbCrLf & vbCrLf & constEmailFooter
        if osmtp.SendMail = 0 then
            submitformemail = formText(constcontactformsuccesstext,"","divFormTitle")
        else
            submitformemail = formText(oSmtp.GetLastErrDescription(),"","divFormTitle errorText")
        end if
		set osmtp = nothing
    end function
%>
<%
    function submitformdb()
        dim placeholder:placeholder = generateRandomValue(50)
        dim idholder
        set conn = server.CreateObject("adodb.connection")
            conn.open databaseConnectionString
        set rs = server.CreateObject("adodb.recordset")
            sql = "SELECT * FROM Company;"
            rs.open sql,conn,3,3    
                rs.addnew
                rs.fields("CompanyName") = placeholder
                rs.update
            rs.close
            sql = "SELECT * FROM Contact;"
            rs.open sql,conn,3,3
                rs.addnew
                rs.fields("ContactFirstName") = placeholder
                rs.update
            rs.close
            sql = "SELECT * FROM Company WHERE Company.CompanyName = '" & placeholder & "';"
            rs.open sql,conn,3,3
                rs.fields("CompanyName") = varray(9)
                rs.fields("CompanyStreet") = varray(10)
                rs.fields("CompanyCity") = varray(11)
                rs.fields("CompanyState") = varray(12)
                rs.fields("CompanyZip") = varray(13)
                rs.fields("CompanyDUNS") = varray(16)
                rs.fields("CompanyPhone") = varray(14)
                rs.fields("CompanyFax") = varray(15)
                idholder = rs.fields("CompanyID")
                rs.update
            rs.close
            sql = "SELECT * FROM Contact WHERE Contact.ContactFirstName = '" & placeholder & "';"
            rs.open sql,conn,3,3
                rs.fields("ContactFirstName") = varray(0)
                rs.fields("ContactLastName") = varray(1)
                rs.fields("ContactTypeID") = varray(2)
                if varray(2) = "9" then
                rs.fields("ContactTypeOther") = varray(3)
                end if
                rs.fields("ContactEmail") = varray(6)
                rs.fields("ContactPassword") = varray(7)
                rs.fields("CompanyID") = idholder
                rs.fields("ContactPhone") = varray(4)
                rs.fields("ContactFax") = varray(5)
                idholder = rs.fields("ContactID")
                rs.update
            rs.close
            sql = "SELECT * FROM Tenant;"
            rs.open sql,conn,3,3
                rs.addnew
                rs.fields("TenantName") = varray(17)
                rs.fields("TenantTypeID") = varray(20)
                rs.fields("TenantContactName") = varray(18)
                rs.fields("SquareFootageID") = varray(19)
                rs.fields("ContactID") = idholder
                rs.fields("TenantDescription") = varray(21)
                rs.update
            rs.close
        set rs = nothing
        set conn = nothing
    end function
%>
<%
'   FORM GENERATION FUNCTION KEY
'   gInput(ftype,fname,fclass,fvalue,flength)
'   gSelectDb(fname,fclass,ftable,foption,fvalue)
'   formText(ftext,fclass)
%>
