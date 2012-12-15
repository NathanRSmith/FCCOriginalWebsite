<%
    if lcase(request.ServerVariables("SCRIPT_NAME")) = "/frames/frame_nformconfirm_contactus.asp" or request.ServerVariables("SCRIPT_NAME") = "/frames/frame_nformaction_contactus.asp" then
        response.CacheControl = "no-cache"
    end if
   
    dim contactfirstname,contactlastname,contactcompany,contactposition,contactcity,contactstate,contactzip,contactcountry,contactphone,contactfax,contactemail,contactperson,comments
    dim validateskiplist:validateskiplist = ",[contactperson]person to contact,"
    dim validatedblookup:validatedblookup = ""
    dim databaseConnectionString:databaseConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Server.MapPath ("/db/nFORMS.mdb")

    dim contactarray:contactarray = Array("General Inquiry - Faith Community Church","Randy Smith - Lead Pastor","Mark Gardner - Associate Pastor","Shannon Runnels - Administrative Assistant")
    dim contactvaluearray()
    for num = 0 to ubound(contactarray)
        redim preserve contactvaluearray(num):contactvaluearray(num) = num
    next
    dim contactemailarray:contactemailarray = Array("faithcommunity@faithcommunityonline.us","randysmith@faithcommunityonline.us","markgardner@faithcommunityonline.us","shannonrunnels@faithcommunityonline.us")
    
        contactfirstname = lv("contactfirstname")
        contactlastname = lv("contactlastname")
        contactcompany = lv("contactcompany")
        contactposition = lv("contactposition")
        contactcity = lv("contactcity")
        contactstate = lv("contactstate")
        contactzip = lv("contactzip")
        contactcountry = lv("contactcountry")
        contactphone = lv("contactphone")
        contactfax = lv("contactfax")
        contactemail = lv("contactemail")
        contactperson = bv("contactperson")
        comments = lv("comments")
 
    dim varray:varray = array(contactfirstname,contactlastname,contactcompany,contactposition,contactcity,contactstate,contactzip,contactcountry,contactphone,contactfax,contactemail,contactperson,comments)
    dim vlist:vlist = "contactfirstname,contactlastname,contactcompany,contactposition,contactcity,contactstate,contactzip,contactcountry,contactphone,contactfax,contactemail,contactperson,comments"
    dim commonlist:commonlist = "First Name,Last Name,Company,Position,City,State,Zip Code,Country,Phone Number,Fax Number,Email Address,Person to Contact,Comments"
    dim ImageValidationArray:ImageValidationArray = Array("5671","5407","4058","4255","1822","6325","6786","8406","2751","2576")
%>