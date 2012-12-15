<%
'****************************************************************************************************************************************
'*	SECTION ONE		SECTION ONE		SECTION ONE		SECTION ONE		SECTION ONE		SECTION ONE		SECTION ONE		SECTION ONE			*
'*																																		*
'*	GATHER PAGE AND ACTION INFORMATION																									*
'*	Public Sub getPA(field,fields)																										*
'*	Public Sub getimageinfo(strpage)																									*
'*																																		*
'****************************************************************************************************************************************
	'------------------------------------------------------------------------------------------------------------------------------------
	'	getPA(field,fields)		getPA(field,fields)		getPA(field,fields)		getPA(field,fields)		getPA(field,fields)		
	'	
	'		
	'************************************************************************************************************************************

Public Sub getPA(field,fields)
	Select Case field
		Case "strpage"																				'//gather page IDentifier			//
				If Not IsNull (strPage) Then
					If Len (strPage) Then
						strpage = strpage
					ElseIf Not Len (strPage) Then
						strpage = Null																'//set page variable to NULL		//
					End If		
				End If					
				If IsNull (strpage) Then				
						If	Len(Request.QueryString(field)) Then									'//if the querystring value exists	//
							strpage	= request.querystring(field)									'//set the page variable to the		//
																									'//querystring value				//
					ElseIf	Len(Request.QueryString(fields)) Then									'//repeat with a second field if	//
							strpage = request.QueryString(fields)									'//the first does not exist			//
					End If					
				End If
				
				If Not IsNull (strpage) Then														'//if the page variable was assigned//
					Call getimageinfo(strpage)														'//a value then continue with the   //
																									'//page generation					//
				ElseIf IsNull (strpage) Then														'//if the variable was Null then	//
					strpage  = Null																	'//set the proper system variables  //
					pgselect = False																'//and return control to the client //
					timg = -1					
				End If	
	
		Case "straction"																			'//gather the action IDentifier		//
			If Not IsNull (pgaction) Then															'//if a value has already been		//
				If Len (pgaction) Then																'//assigned to the pgaction variable//
					pgaction = pgaction																'//check the length and continue	//
				ElseIf Not Len (pgaction) Then
					pgaction = Null																	'//set action variable to NULL	if	//
				End If																				'//the value had a 'false' length	//
			End If																					
			If	Len(Request.QueryString(field)) Then												'//check the querystring for an		//
				pgaction = request.querystring(field)												'//appropriate action value			//
																									'//and assign the value to pgaction //
				If IsNull (pgaction) Then															'//if the value did not exist return//
					pgaction = null																	'//control to the client to select  //
				End If																				'//an action						//
				
			End If
			pgselect = True																			'//tell the system that a page		//
																									'//has been selected				//
		Case Else
		
	End Select
End Sub
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub getimageinfo(strpage)	Public Sub getimageinfo(strpage)	Public Sub getimageinfo(strpage)	
	'	
	'	
	'************************************************************************************************************************************
	'******************************************************************************************GET IMAGE INFORMATION FOR THE CURRENT PAGE
																					
Public Sub getimageinfo(strpage)
	SQL=	"SELECT Image_REL_Page.*, Image_Inf.Image FROM Image_Inf "							'//set the SQL commands to get page
	SQL=SQL&"INNER JOIN Image_REL_Page ON Image_Inf.ImageID = Image_REL_Page.ImageID "			'//specific image information 
	SQL=SQL&"WHERE (((Image_REL_Page.PageID)='"&strpage&"'));"		
	Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)										'//open the recordset and connection
		If Not RS.eof Then																		'//continue if relational images exist
			timg = Cint(RS.RecordCount-1)														'//[ubound]variable total images minus 1
			ReDim arimg(timg)																	'//create array arimg(total images)
			ReDim arrel(2,timg)																	'//three dimensional array to replace 
																								'//arimg() in revision
			For	i =	0 To timg																	'//For...Next to gather image ID's for all
				arimg(i)  =RS("imageID")														'//images related to the current page
				arrel(0,i)=RS("REL_Image_PageID")												'//relational ID	[for revision]
				arrel(1,i)=RS("ImageID")														'//image ID			[for revision]
				arrel(2,i)=RS("PageID")															'//page ID			[for revision]
				If Not RS.EOF Then																'//move through the recordset one row if
					RS.movenext																	'//the last row has not yet been reached
				End If 
			Next
		Else																					'//IF THERE ARE NO PAGE RELATED IMAGES
			timg = -1																			'//set the total images variable to -1
			ReDim arimg(0)																		'//set one dimensional empty array
			arimg(0)=""																			'//create an empty array value to 
		End If																					'//deter potential errors
	Call ocon("c","","","")																		'//close the recordset and connection
	Call pginformation(strpage)																	'//get page\image parameters			
	Call getPA("straction",null)																'//check for selected action and continue
End Sub

'****************************************************************************************************************************************
'*	SECTION TWO		SECTION TWO		SECTION TWO		SECTION TWO		SECTION TWO		SECTION TWO		SECTION TWO		SECTION TWO			*
'*																																		*
'*	DYNAMIC PAGE CREATION ONE																											*
'*	Public Sub iHeader()																												*
'*	Public Sub ReadFile(rfile)																											*
'*	Public Sub BodyBuilder(pgaction)																									*
'*	Public Sub loadmod(action)																											*
'*																																		*
'****************************************************************************************************************************************
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub iHeader()	Public Sub iHeader()	Public Sub iHeader()	Public Sub iHeader()	Public Sub iHeader()	
	'	
	'	
	'************************************************************************************************************************************
	'*****************************************************************Append Generic HTML + Javascript + OnLoad Functions to ihtml String
Public Sub iHeader()
	ihtml = ihtml & "<HTML>" & vbCrLf
	ihtml = ihtml & "<HEAD>" & vbCrLf
	ihtml = ihtml & "<TITLE>J&F Electric Image Management</TITLE>" & vbCrLf
	ihtml = ihtml & "<STYLE TYPE=""text/css"">" & vbCrLf
		ihtml = ihtml & ".formstyle {color: #ffffff; font-weight: bold; background-color: #006699; border-color: #006699;}" & vbCrLf
		ihtml = ihtml & "a:visited	{color: white;}" & vbCrLf
		ihtml = ihtml & " .nu		{text-decoration: none;	font-family: sans-serif; font-size:	10pt;}" & vbCrLf
		ihtml = ihtml & ".jfbutton	{color:white;font-weight:bold;background-color:maroon;border-WIDTH:3px;border-color:white;}" & vbCrLf
	ihtml = ihtml & "</STYLE>" & vbCrLf			
	'server.Execute("..\include\jTools.asp")	
	Response.Write (ihtml)
	ihtml = ""										
	%>
	<!--#include file="jtools.asp"-->
	<%
	ihtml = ihtml & "</HEAD>" & vbCrLf
End Sub

	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub ReadFile(rfile)	Public Sub ReadFile(rfile)	Public Sub ReadFile(rfile)	Public Sub ReadFile(rfile)	
	'	
	'	
	'************************************************************************************************************************************
	'******************************************************Uses the FileSystem Object to 'Read' Lines of Javascript into the ihtml String
Public Sub ReadFile(rfile)
	Const ForReading = 1																		'//define the reading constant			//
	Dim objifilejs
	Dim strifilejs
	Call getRoot("file",rfile)																	'//finds root server folder				//
	jspath  = FSO.BuildPath (tmpspec, rfile)													'//Builds path to the file				//
	Set objifilejs = fso.OpenTextFile(jspath, ForReading, False)								'//set the 'Read' File [i.e. jTools.js] //
		Do While NOT objifilejs.AtEndOfStream													'//'Read' each line of the file to the  //
			strifilejs = strifilejs & objifilejs.Readline & vbCrLf								'//temporary variable strifilejs		//
		Loop												
	objifilejs.Close																			'//close the textfile
	ihtml = ihtml & strifilejs																	'//append the temp string to the ihtml
	Set objifilejs = Nothing																	'//string
End Sub

	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub BodyBuilder(pgaction)	Public Sub BodyBuilder(pgaction)	Public Sub BodyBuilder(pgaction)	
	'	
	'	
	'************************************************************************************************************************************
Public Sub BodyBuilder(pgaction)																'//OnLoad Code For Various Pages
	Select Case pgaction															
		Case "Modify Placement"
				Call loadmod(pgaction)															'//Creates dynamic onload script[MODIFY]//
		Case "Add Image"
				%>ONLOAD="forcethumbs('<%=pginfo(1)%>');"<%										'//Creates check|disable Onload for[ADD]//
		Case Else																				'//if the parameters Force Thumbnails	//
				ihtml = ihtml 
	End Select
End Sub

	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub loadmod(action)	Public Sub loadmod(action)	Public Sub loadmod(action)	Public Sub loadmod(action)	
	'	
	'	
	'************************************************************************************************************************************
Public Sub loadmod(action)
	Dim rget,looper,intprei
		SQL="SELECT ImageOrder.REL_Image_PageID, ImageOrder.ImageOrder, Image_Inf.* FROM (Image_Inf INNER JOIN Image_REL_Page ON "
		SQL=SQL&"Image_Inf.ImageID = Image_REL_Page.ImageID) INNER JOIN ImageOrder ON Image_REL_Page.REL_Image_PageID = "
		SQL=SQL&"ImageOrder.REL_Image_PageID WHERE Image_REL_Page.PageID='"&strpage&"' ORDER BY ImageOrder;"
		Select Case action
		Case "Modify Placement"
			If pginfo(0)>=1 Then
				Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
					If Not RS.eof Then
						rget = RS.getrows(-1)
						intprei = rs.recordcount-1
					Else
						intprei = -1
					End If
				Call oCon("c",null,null,null)
					If isarray(rget) Then
						%> onload="<%
						For i = 0 TO intprei
							For ii = 0 TO timg
								IF int(rget(2,i)) = int(arrel(1,ii)) THEN
								%>intimgorder(<%=ii+1%>,'<%=rget(3,i)%>',<%=pginfo(0)%>,'',<%=rget(2,i)%>,<%=rget(0,i)%>,<%=rget(1,i)%>);<%
							End If
							Next
						Next
						ihtml = ihtml & """"
					End If 
			Else
			%>
			onload="alert('This page is not configured for image management. Please contact the system administrator.')"
			<%
			End If
		Case "displayed"
			Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
				If RS.EOF Then
					displayed = 0
				Elseif RS.Recordcount = -1 Then
					displayed = 0
				Elseif Rs.Recordcount > -1 Then
					displayed = RS.Recordcount
				End If
			Call oCon("c",null,null,null)
			
		Case Else
		
	End Select
End Sub

'****************************************************************************************************************************************
'*	SECTION THREE	SECTION THREE	SECTION THREE	SECTION THREE	SECTION THREE	SECTION THREE	SECTION THREE	SECTION THREE		*
'*																																		*
'*	Public Sub DynamicPage(pgaction)																									*
'*	Public Sub uploadform(strpage,pgaction,cont)																						*
'*	Public Sub tophtml(Line)																											*
'*	Public Sub createthumbs(cont)																										*
'*	Public Sub createinteractive(oc)																									*
'*	Public Sub PickaPage(doit)																											*
'*																																		*
'****************************************************************************************************************************************
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub DynamicPage(pgaction)	Public Sub DynamicPage(pgaction)	Public Sub DynamicPage(pgaction)	
	'	
	'	
	'************************************************************************************************************************************
Public Sub DynamicPage(pgaction)
	Select Case	pgaction
		Case "Add Image"																		'//Create HTML for the Add Form			//
			If timg+1 < pginfo(4) AND timg >=0 Then												'//Call Normal Form						//
				Call uploadform(strpage,pgaction,"")											
			Elseif timg >= 0 Then																'//User has reached their upload		//
				Call uploadform(strpage,pgaction,"LIMIT")										'//limit message is displayed			//
			Elseif timg < 0 Then																
				Call uploadform(strpage,pgaction,"NOIMAGE")										'//No page related images exist so		//
			End If																				'//the page is loaded without thumbnails//
			
		Case "Modify Placement"
			Call tophtml("<B>Modify the Order With Which Your Images Appear.</B>")				'//Generic HTML appended to ihtml string//
			If timg >= 0 Then																	'//If Images exist						//
				Call CreateInteractive("mod")													'//append javascript powered thumbnails //
			Elseif timg < 0 Then																'//If Images Do Not Exist				//
				Call CreateInteractive("NOIMAGE")												'//append message to imsg variable		//
			End If
			
		Case "Remove Image"	
			Call tophtml("<B>Click the Images Below to Remove. You May Select Multiple Images.</B>")	  '//insert generic top html	//
			If timg >= 0 Then																	'//If Images exist						//
				Call CreateInteractive("rem")													'//insert javascript powered thumbnails //
			Elseif timg < 0 Then																'//If Images Do Not Exist				//
				Call CreateInteractive("NOIMAGE")												'//append message to imsg variable		//
			End If
			
		Case Else
			If Not pgselect	Then																'//If a page has not been selected		//
				Call PickaPage("no")															'//system defaults to page selection	//
			Else
				Call PickaPage("yes")															'//The page has been selected so the	//
				If timg >= 0 Then																'//system loads the select action page	//
					Call CreateThumbs("IMAGE")													'//If the image recordset exists then	//
				Elseif timg < 0 Then															'//load normally. If it does not exist	//
					Call CreateThumbs("NOIMAGE")												'//then return a message.				//
				End If
			End If																				
	End Select
End Sub

	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub uploadform(strpage,pgaction,cont)	Public Sub uploadform(strpage,pgaction,cont)	
	'	
	'	
	'************************************************************************************************************************************
	'************************************************************************************************Creates HTML for ADD IMAGE qualifier
Public Sub uploadform(strpage,pgaction,cont)
If cont <> "LIMIT" then
	Call tophtml("<b>Upload an Image to the Server.</b>")
Else
	Call tophtml("LIMIT")
End If
If cont <> "LIMIT" Then
	ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
		ihtml = ihtml & "<TD ALIGN=""left"">"
		ihtml = ihtml & "Alt Text"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""3"">"
		ihtml = ihtml & "<INPUT TYPE=""text"" NAME=""stralt"" ID=""stralt"" STYLE=""WIDTH:400;background-color:maroon;color:white;border:white 1px dashed;font-weight:bold;"">"
		ihtml = ihtml & "</A>"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
		ihtml = ihtml & "<TD ALIGN=""left"">"
		ihtml = ihtml & "Image"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""3"">"
		ihtml = ihtml & "<INPUT TYPE=""file"" ID=""imagemgtup"" NAME=""imagemgtup"" STYLE=""WIDTH:400;background-color:maroon;color:white;border:white 1px dashed;font-weight:bold;"">"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
		ihtml = ihtml & "<TD>"
		ihtml = ihtml & "Create Thumbnail"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD Colspan=""3"">"
		ihtml = ihtml & "<INPUT TYPE=""checkbox"" NAME=""boothumb"" ID=""boothumb"" STYLE=""background-color:maroon;color:white;border:white 1px dashed;font-weight:bold;"">"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
		ihtml = ihtml & "<TD>"  & vbCrLf
		ihtml = ihtml & "Crop Method"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD align=""left""  Colspan=""3"">"  & vbCrLf
		ihtml = ihtml & "<SELECT Name=""icrop"" ID=""icrop"" STYLE=""width:150;background-color:maroon;color:white;border:white 1px dashed;font-weight:bold;"">"
		ihtml = ihtml & "<OPTION Value=""topleft"" selected>[Top | Left] Crop</OPTION>"
		ihtml = ihtml & "<OPTION Value=""center"">Center Crop</OPTION>"
		ihtml = ihtml & "<OPTION Value=""botright"">[Bottom | Right] Crop</OPTION>"
		ihtml = ihtml & "</SELECT>"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
		ihtml = ihtml & "<TD ALIGN=""left"">"
		ihtml = ihtml & "<INPUT TYPE=""HIDDEN"" NAME=""hideme"" ID=""hideme"" VALUE=""Add Image"">"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""3"">"
		ihtml = ihtml & "&nbsp;"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR>" & vbCrLf
		ihtml = ihtml & "<TD>"
		ihtml = ihtml & "&nbsp;"
		ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""3"">"
			If cont <> "LIMIT" Then
			ihtml = ihtml & "<INPUT type=""submit"" value="""&pgaction&""" name=""Submit1"" class=""jfbutton"" STYLE=""WIDTH:150;"">"
			Elseif cont = "LIMIT" Then
			ihtml = ihtml & "<INPUT type=""submit"" value="""&pgaction&""" name=""Submit1"" class=""jfbutton"" STYLE=""WIDTH:150;"" disabled>"
			End If
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR height=""10"">" & vbCrLf
		ihtml = ihtml & "<TD	colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR height=""3"" bgcolor=""maroon"">" & vbCrLf
		ihtml = ihtml & "<TD colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	Else
	ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""4"">"
		ihtml = ihtml & "<b>You Have Met or Exceeded Your Maximum Image Uploads.<br>Please Remove One or More Images to Unlock the Add Images Feature.</b>"
		ihtml = ihtml & "</A>"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	End If
	If cont <> "NOIMAGE" Then
		Call CreateThumbs("IMAGE")																	'//Creates the thumbnailed images	//
	Elseif cont = "NOIMAGE" Then
		Call CreateThumbs("NOIMAGE")	
	End If
End Sub
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub tophtml(Line)	Public Sub tophtml(Line)	Public Sub tophtml(Line)	Public Sub tophtml(Line)		
	'	
	'	
	'************************************************************************************************************************************
Public Sub tophtml(Line)
	ihtml = ihtml & "<BODY bgcolor=""#330000"" TEXT=""White"""
	Response.Write (ihtml)
	ihtml = ""
	If Line <> "LIMIT" Then
		Call BodyBuilder(pgaction)																	'//Builds Action specific Body OnLoad	
	Else
		Line = "Image Upload Functionality Temporarily Disabled"
	End If
	ihtml = ihtml & ">" & vbCrLf
	Response.Write (ihtml)
	ihtml = ""
	If line <> "initial" Then
	%>
	<form name="form1" onsubmit="return finalaction('imagemgtaction.asp','GET','<%=pgaction%>',<%=pginfo(0)%>);" enctype="multipart/form-data">
	<%
	ihtml = ihtml & "<table>" & vbCrLf
	ihtml = ihtml &	"<TR valign=""top"">" & vbCrLf
		ihtml = ihtml & "<TD align=""left"" colspan=""2"">"
			ihtml = ihtml & "<a href=""imagemgt.asp?strpage="&strpage&""">"
			ihtml = ihtml & "<B>Return TO the Image Management Menu.</B>"
			ihtml = ihtml & "</a>"
			ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "<TD align=""right"" colspan=""2"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR bgcolor=""white"" STYLE=""height:3;"">" & vbCrLf
		ihtml = ihtml & "<TD colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR valign=""top"">" & vbCrLf
		ihtml = ihtml & "<TD align=""left"" colspan=""4"">"
		ihtml = ihtml & Line
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR STYLE=""height:10;"">" & vbCrLf
		ihtml = ihtml & "<TD	colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	Else
	
	End If
End Sub
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub createthumbs(cont)	Public Sub createthumbs(cont)	Public Sub createthumbs(cont)	Public Sub createthumbs(cont)	
	'	
	'	
	'************************************************************************************************************************************
	'*************************************Creates the Thumbnails For the Non-Interactive Pages(add/ELSE) ... Combine With Other SUB Later
Public Sub createthumbs(cont)
If cont <> "NOIMAGE" Then					
	For	i =	0 To Ubound(arimg) Step 3
		SQL	= "SELECT *	FROM Image_inf WHERE Image_inf.ImageID = " & arimg(i)
		For ii = 1 To 3
			If ii+i <= Ubound(arimg) Then
			SQL	= SQL &	" or image_inf.imageID = " & arimg(ii+i)
			End	If
		Next
			SQL	= SQL &	";"
			Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
			If Not Rs.EOF Then
			ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD STYLE=""WIDTH:150;align:right;"">"
				If i = 0 Then
				ihtml = ihtml & "<B>"
				ihtml = ihtml & "Image:"
				ihtml = ihtml & "</B>"
				End If
			ihtml = ihtml & "</TD>" & vbCrLf
			For ii = 1 To 3
			ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
				If (ii+i)-1 <= UBOUND(arimg) THEN	
				ihtml = ihtml & RS("image")
				Else
				ihtml = ihtml & "&nbsp;"
				End if 
			ihtml = ihtml & "</TD>" & vbCrLf
			If Not RS.eof Then
				RS.movenext
			End If 
			Next
			RS.movefirst
			ihtml = ihtml & "</TR>"  & vbCrLf
			ihtml = ihtml & "<TR height="""&pginfo(3)*1.15&""">" & vbCrLf
				ihtml = ihtml & "<TD STYLE=""WIDTH:150"">"
				ihtml = ihtml & "&nbsp;"
				ihtml = ihtml & "</TD>" & vbCrLf
			FOR ii = 1 TO 3
			ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
				IF (ii+i)-1 <= UBOUND(arimg) THEN	
				ithumb = getthumb(rs("image"))
				ihtml = ihtml & "<img STYLE=""border:white 2px SOLID;"" height="""& pginfo(3) &""" alt=""" & RS("stralt") & """ src=""../images/" & ithumb & """>"
				ELSE
				ihtml = ihtml & "&nbsp;"
				End If
			ihtml = ihtml &  "</TD>" & vbCrLf
				IF NOT RS.eof THEN
				RS.movenext
				end IF
			NEXT
		RS.movefirst
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR height="""&pginfo(3)*.1&""">" & vbCrLf
		ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
	'		If i = 0 Then
	'		ihtml = ihtml & "<B>"
	'		ihtml = ihtml & "Thumbnailed:"
	'		ihtml = ihtml & "</B>"
	'		End If
		ihtml = ihtml & "</TD>" & vbCrLf
		For ii = 1 To 3
		ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
			If (ii+i)-1 <= Ubound(arimg) Then	
			ihtml = ihtml & "Thumbnailed"
			Else
			ihtml = ihtml & "&nbsp;"
			End If 
		ihtml = ihtml & "</TD>"
			If Not RS.eof Then
			RS.movenext
			End If
		Next
		RS.movefirst
		ihtml = ihtml & "</TR>"
		ihtml = ihtml & "<TR height="""&pginfo(3)*.1&""">" & vbCrLf
		ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
			If i = 0 Then
			ihtml = ihtml & ""
			ihtml = ihtml & "</TD>" & vbCrLf
			End If
		FOR ii = 1 TO 3
		ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
			IF (ii+i)-1 <= UBOUND(arimg) THEN	
			ihtml = ihtml & "&nbsp;"
			ELSE
			ihtml = ihtml & "&nbsp;"
			End If
		ihtml = ihtml &  "</TD>" & vbCrLf
			If Not RS.eof Then
			RS.movenext
			End If
		Next
		RS.movefirst
		ihtml = ihtml & "</TR>" & vbCrLf
		If i >= Ubound(arimg)-2 then
		ihtml = ihtml & "<TR align=""right"">" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""4"">"
				ihtml = ihtml & "<a href=""LOGOUT.ASP"" STYLE=""FONT-WEIGHT:BOLD;COLOR:WHITE;"">LOGOUT</A>"
				ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
			Else
		End If
		End If
		Call oCon("c","","","")
		Next
ElseIf cont = "NOIMAGE" Then			
	ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
		ihtml = ihtml & "<TD COLSPAN=""4"">"
		ihtml = ihtml & "There Are No Images To Display. Please Begin By Using the Add Images Feature."
		ihtml = ihtml & "</A>"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR height=""10"">" & vbCrLf
		ihtml = ihtml & "<TD colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR height=""3"" bgcolor=""maroon"">" & vbCrLf
		ihtml = ihtml & "<TD	colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	ihtml = ihtml & "<TR height=""10"">" & vbCrLf
		ihtml = ihtml & "<TD align=""right"" colspan=""4"">"
			ihtml = ihtml & "<a href=""LOGOUT.ASP"" STYLE=""FONT-WEIGHT:BOLD;COLOR:WHITE;"">LOGOUT</A>"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
End If
End Sub
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub createinteractive(oc)	Public Sub createinteractive(oc)	Public Sub createinteractive(oc)	
	'	
	'	
	'************************************************************************************************************************************
	'*****************************************************************************CREATES THE INTERACTIVE THUMBNAILS IN MOD AND REM PAGES
Public Sub createinteractive(oc)
	If oc <> "NOIMAGE"	Then		'<---- If Images Exist | Create Dynamic Content

		If oc = "mod"	Then
			irows = pginfo(0)-1		'<---- [Displayable Image Slots] - 1
		Else
			irows = timg			'<---- [Total Images For Managed Page] - 1
		End If
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""4"" height=""10"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		
		If oc = "mod" Then			'<---- If Modifying Image Order Continue Here
		
			ihtml = ihtml & "<TR>" & vbCrLf		
				ihtml = ihtml & "<TD colspan=""4"" height=""3"" bgcolor=""maroon"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR>" & vbCrLf
				ihtml = ihtml & "<TD	colspan=""4"" height=""10"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR>" & vbCrLf
			
									'<---- Modify First Tier [ORDER IMAGES] Thumbnail Creation Begins Here
									
			For i = 0 To (pginfo(0)-1) Step 3		
				ihtml = ihtml & "<TR STYLE=""height:"&pginfo(3)*.9&";"">" & vbCrLf
					ihtml = ihtml & "<TD valign=""top"" STYLE=""WIDTH:150;align:right;"">"
						If i = 0 Then
							ihtml = ihtml & "<b>"
							ihtml = ihtml & "Order:"
							ihtml = ihtml & "</b>"
						Else
							ihtml = ihtml & "&nbsp;"
						End If
					ihtml = ihtml & "</TD>" & vbCrLf
					For ii = 1 To 3
						ihtml = ihtml & "<TD	ALIGN=""center"" STYLE=""WIDTH:150;"">"
						If i+ii<=pginfo(0) Then
							ihtml = ihtml & "<a href=""#"">"
							Response.Write (ihtml)
							ihtml = ""
							%> 
							<IMG name="" border="0" STYLE="border-color:white;"	height="0" ID="rimage<%=i+ii%>" onclick="intimgorder(<%=i+ii%>,'',<%=pginfo(0)%>,'rem','','',false)">
							<%
							ihtml = ihtml & "</a>"
						Else
							ihtml = ihtml & "&nbsp;"
						End If
						ihtml = ihtml & "</TD>" & vbCrLf
					Next
					ihtml = ihtml & "</TR>" & vbCrLf
				ihtml = ihtml & "<TR>" & vbCrLf
					ihtml = ihtml & "<TD STYLE=""WIDTH:150;align:right;"">"
					ihtml = ihtml & "<B>"
					ihtml = ihtml & "&nbsp;"
					ihtml = ihtml & "</B>"
					ihtml = ihtml & "</TD>" & vbCrLf
					For ii = 1 To 3
						ihtml = ihtml & "<TD	ALIGN=""center"" STYLE=""WIDTH:150;"">"
							If i+ii<=pginfo(0) Then
								ihtml = ihtml & "["&i+ii&"]"
							Else
								ihtml = ihtml & "&nbsp;"
							End If 
						ihtml = ihtml & "</TD>" & vbCrLf
					Next
				ihtml = ihtml & "</TR>" & vbCrLf
			Next
									
									'<---- Modify First Tier [ORDER IMAGES] Thumbnail Creation Ends Here
									
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""4"" height=""10"">"
			ihtml = ihtml & "&nbsp;"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		
		End If

		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"" height=""3"" bgcolor=""maroon"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"" height=""10"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf

									'<---- Common Dynamic Second Tier Thumbnail Creation Begins Here

		For	i =	0 To Ubound(arimg) Step	3
			SQL = ""
			SQL	=		"SELECT Image_Inf.*, Image_REL_Page.REL_Image_PageID "
			SQL = SQL & "FROM Image_Inf INNER JOIN Image_REL_Page ON Image_Inf.ImageID = "
			SQL = SQL & "Image_REL_Page.ImageID WHERE Image_Inf.ImageID = " & arimg(i)
			For ii=1 To 2
				If i + ii <= Ubound(arimg) Then
					SQL	= SQL &	" OR Image_Inf.ImageID = " & arimg(i+ii)
				End If
			Next	
			SQL = SQL & ";"
			Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
				moveme = rs.Bookmark
				ReDim ocx(timg)
				ihtml = ihtml & "<TR HEIGHT="""&pginfo(3)*.9&""">" & vbCrLf
					ihtml = ihtml & "<TD VALIGN=""top"" STYLE=""WIDTH:150;align:right;"">"
						If i = 0 Then
							ihtml = ihtml & "<B>"
							ihtml = ihtml & "Image:"
							ihtml = ihtml & "</B>"
						Else
							ihtml = ihtml & "&nbsp;"
						End If
					ihtml = ihtml & "</TD>" & vbCrLf
					For ii=1 To 3
						ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
						If i+ii <= Ubound(arimg)+1 Then
							Do While z <= Ubound(arimg)
								If arrel(1,z)=RS("imageID") Then
									Exit Do
								Else
									z = z+1
								End If
							Loop
							ihtml = ihtml & "<a href=""#"">" & VbCrLf
							x = (i + ii)-1						

						'********************************************************************
						' The loops below are a quick code to check the two dimensional		*
						' array FOR the image order creating the images accordingly			*
						' Later, code this correctly...shouldn't hurt anything, but it		*
						' is less efficient may pose problems later?						*
						'********************************************************************
						ithumb = getthumb(rs("image"))
						SELECT CASE oc
							CASE "mod"
								IF x<=timg THEN
									Response.Write (ihtml)
									ihtml = ""
									%>
									<img height="<%=pginfo(3)*.75%>" width="<%=pginfo(2)*.75%>" onclick="intimgorder(<%=i+ii%>,'<%=ithumb%>',<%=pginfo(0)%>,'',<%=RS("imageID")%>,<%=arrel(0,z)%>,false);" STYLE="border:maroon 2px solid" ID="image<%=i+ii%>" src="../images/<%=ithumb%>">
									<%
									subone="Modify Placement"
									x = pginfo(0)-1
									plus = 1
								ELSE
									
								END IF
							CASE "rem"
								IF x<=ubound(arimg) THEN
									Response.Write (ihtml)
									ihtml = ""
									%>
									<img border="0" height="<%=pginfo(3)*.75%>" width="<%=pginfo(2)*.75%>"STYLE="border:maroon 2px solid;" ID="image<%=(ii+i)-1%>" NAME="<%=RS("REL_Image_PageID")%>" src="../images/<%=ithumb%>" onclick="intimgselect(<%=(ii+i)-1%>,<%=UBOUND(arimg)%>);">
									<%
									subone="Remove Image(s)"
									x = x
									plus = 0
								ELSE
									
								END IF
							CASE ELSE
								IF x<=ubound(arimg) THEN
									Response.Write (ihtml)
									ihtml = ""
									%>
									<img border="0" height="<%=pginfo(3)*.75%>" width="<%=pginfo(2)*.75%>"STYLE="border:2px solid maroon;" ID="image<%=(ii+i)-1%>" src="../images/<%=ithumb%>)">
									<%
									subone="Continue"
									x = x
									plus = 0
								ELSE
								
								END IF
					END SELECT
					
					'<---- End Array Check Loops Here
				ihtml = ihtml & "</a>"
				Else
					ihtml = ihtml & "&nbsp;"
				End If
				ihtml = ihtml & "</TD>" & vbCrLf
				IF NOT RS.eof THEN
					RS.movenext
				END IF 
			Next
			ihtml = ihtml & "</TR>" & vbCrLf
			RS.bookmark = moveme
			z = 0
						
			ihtml = ihtml & "<TR height=""10"">" & vbCrLf
				ihtml = ihtml & "<TD STYLE=""WIDTH:150"">"
				ihtml = ihtml & "&nbsp;"
				ihtml = ihtml & "</TD>" & vbCrLf
				For ii=1 To 3
				ihtml = ihtml & "<TD STYLE=""WIDTH:150;"">"
					If i+ii <= Ubound(arimg)+1 Then
						ihtml = ihtml & RS("image")
						If Not RS.eof Then
							RS.movenext
						End If 
					Else
						ihtml = ihtml & "&nbsp;"
					End If
				ihtml = ihtml & "</TD>" & vbCrLf
				Next
			ihtml = ihtml & "</TR>" & vbCrLf
			RS.bookmark = moveme
		Call oCon("c","","","")
		Next
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"" height=""10"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"" height=""3"" bgcolor=""maroon"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"" height=""10"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR valign=""bottom"">" & vbCrLf
			ihtml = ihtml & "<TD STYLE=""WIDTH:150"">"
			ihtml = ihtml & "&nbsp;"
			ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "<TD colspan=""3"" cellpad=""0"" cellspacing=""0"" valign=""bottom"" STYLE=""WIDTH:450;"">"	
		For	i =	0 To irows			'<---- CREATES THE TEXT FIELDS FOR MODIFY AND REMOVE FUNCTIONALITIES
			If oc = "mod" Then
				ihtml = ihtml & "<INPUT type=""text"" ID=""imagemgtup"&i+1&""" "
				ihtml = ihtml & "STYLE=""background-color:maroon;font-size:12pt;"
				ihtml = ihtml & "font-family:courier;text-transform:uppercase;color:white;"
				ihtml = ihtml & "border:white 1 dashed;"" name=""imagemgtup"" STYLE=""WIDTH:450"" readonly>"
				ihtml = ihtml & "<INPUT type=""HIDDEN"" ID=""ximagemgtup"&i+1&""" name=""ximagemgtup"">"
			Else
				ihtml = ihtml & "<INPUT type=""text"" ID=""imagemgtup"&i&""" "
				ihtml = ihtml & "STYLE=""background-color:maroon;font-size:12pt;"
				ihtml = ihtml & "font-family:courier;text-transform:uppercase;color:white;"
				ihtml = ihtml & "border:white 1 dashed;"" name=""imagemgtup"" STYLE=""WIDTH:450"" readonly>"
				ihtml = ihtml & "<INPUT type=""HIDDEN"" ID=""ximagemgtup"&i&""" name=""ximagemgtup"">"
			End If
		Next
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR height=""10"">" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""4"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD STYLE=""WIDTH:150"">"
			ihtml = ihtml & "&nbsp;"
			ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""3"">"
			ihtml = ihtml & "<INPUT type=""HIDDEN"" name=""hideme"" value="""&pginfo(0)&""">"
			ihtml = ihtml & "<INPUT type=""submit"" name=""Submit1""	value="""&subone&""" class=""jfbutton"" STYLE=""WIDTH:450;"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR height=""10"">"
			ihtml = ihtml & "<TD	align=""right"" colspan=""4"">"
			If i+ii >= irows Then
				ihtml = ihtml & "<a href=""LOGOUT.ASP"" STYLE=""FONT-WEIGHT:BOLD;COLOR:WHITE;"">LOGOUT</A>"
				ihtml = ihtml & "</TD>" & vbCrLf
			Else
				ihtml = ihtml & "&nbsp;"
			End If
		ihtml = ihtml & "</TR>" & vbCrLf
									'<---- End Dynamic Thumbnail Creation Here
	Else							'<---- If Images Related to the Current Page Do Not Exist Return This Message
		ihtml = ihtml & "<TR VALIGN=""bottom"">" & vbCrLf
			ihtml = ihtml & "<TD COLSPAN=""4"">"
			ihtml = ihtml & "There Are No Images To Display. Please Begin By Using the Add Images Feature."
			ihtml = ihtml & "</A>"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>"
		ihtml = ihtml & "<TR height=""10"">" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR height=""3"" bgcolor=""maroon"">" & vbCrLf
			ihtml = ihtml & "<TD colspan=""4"">"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
		ihtml = ihtml & "<TR height=""10"">" & vbCrLf
			ihtml = ihtml & "<TD align=""right"" colspan=""4"">"
			ihtml = ihtml & "<a href=""LOGOUT.ASP"" STYLE=""FONT-WEIGHT:BOLD;COLOR:WHITE;"">LOGOUT</A>"
			ihtml = ihtml & "</TD>" & vbCrLf
		ihtml = ihtml & "</TR>" & vbCrLf
	End If
End Sub
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub PickaPage(doit)	Public Sub PickaPage(doit)	Public Sub PickaPage(doit)	Public Sub PickaPage(doit)		
	'	
	'	LOOK HERE
	'************************************************************************************************************************************
Public Sub PickaPage(doit)
	Call tophtml("initial")
	Select Case doit
		Case "no"
			ihtml = ihtml & "<form name=""form1"" method=""get"" action=""imagemgt.asp"" enctype=""multipart/form-data"">" & vbCrLf
			ihtml = ihtml & "<table>" & vbCrLf
			SQL	= "SELECT DISTINCT pagerights.pageID FROM pagerights WHERE pagerights.personnelID =	'" & session("pid")	& "';"
			'ihtml = ihtml & "SESSION PID=" & session("pid") & vbCrLf & "STATE=" & rs.state & vbCrLf & "sql=" & sql & vbCrLf			
			Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
				If Not RS.EOF Then
					xx=RS.recordcount-1
						IF xx =	0 Then
							strpage = rs("PageID")
							Call oCon("c","","","")
							Response.Redirect ("imagemgt.asp?strpage="&strpage)
						End If
					ArrPage = RS.Getrows()
					RS.movelast
				Else
					ReDim ArrPage(0,0)
					ArrPage(0,0) = "Error:Database.EOF - Please Try Again"
					xx = 0
					Call oCon("c","","","")
				End	If
			ihtml = ihtml & "<TR>" & vbCrLf
			ihtml = ihtml & "<TD	colspan=""4"" align=""left"">" & vbCrLf
				ihtml = ihtml & "<select	name=""strpage"">" & vbCrLf
				For	i =	0 to xx
				ihtml = ihtml & "<option	value=""" &	arrpage(0,i) & """>"
				ihtml = ihtml & arrpage(0,i)
				ihtml = ihtml & "</option>" & vbCrLf
				Next
				ihtml = ihtml & "</select>" & vbCrLf
			ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR>" & vbCrLf
				ihtml = ihtml & "<TD	colspan=""4"" height=""15px"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR>" & vbCrLf
				ihtml = ihtml & "<TD colspan=""4"" align=""left"">"
				ihtml = ihtml & "<INPUT type=""submit"" value=""Select Page"" class=""jfbutton"" STYLE=""WIDTH:150;"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "</table>" & vbCrLf
		Case "yes"
			ihtml = ihtml & "<table>" & vbCrLf
			response.Write (ihtml)
			ihtml = ""
			%>
			<form name="form1" method="get" action="imagemgt.asp" onsubmit="return actionform();" enctype="text/html">
			<%
			ihtml = ihtml & "<TR valign=""bottom"">" & vbCrLf
				ihtml = ihtml & "<TD align=""left""	STYLE=""WIDTH:150;"">"
				ihtml = ihtml & "<B STYLE=""font-size:12pt;"">Select An Action: </B>"
					ihtml = ihtml & "&nbsp;&nbsp;"
					ihtml = ihtml & "<INPUT type=""HIDDEN"" name=""straction"" value="""&pgaction&""">"
					ihtml = ihtml & "<INPUT type=""HIDDEN"" name=""strpage"" value="""&strpage&""">"
				ihtml = ihtml & "</TD>"
				ihtml = ihtml & "<TD>"
					Response.Write (ihtml)
					ihtml = ""
					%>
					<INPUT type="submit" value="Add Image" onclick="document.forms[0].straction.value = 'Add Image';" class="jfbutton" STYLE="WIDTH:150;">
					<%
				ihtml = ihtml & "</TD>"
				ihtml = ihtml & "<TD>"
					Response.Write (ihtml)
					ihtml = ""
					%>
					<INPUT type="submit" value="Remove Image" onclick="document.forms[0].straction.value = 'Remove Image';" class="jfbutton" STYLE="WIDTH:150;">
					<%
				ihtml = ihtml & "</TD>"
				ihtml = ihtml & "<TD>"
					Response.Write (ihtml)
					ihtml = ""
					%>
					<INPUT type="submit" value="Modify Placement" onclick="document.forms[0].straction.value =	'Modify Placement';" class="jfbutton" STYLE="WIDTH:150;">
					<%
				ihtml = ihtml & "</TD>"
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR height=""3"" bgcolor=""white"">" & vbCrLf
				ihtml = ihtml & "<TD	colspan=""4"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
			ihtml = ihtml & "<TR height=""10"">" & vbCrLf
				ihtml = ihtml & "<TD	colspan=""4"">"
				ihtml = ihtml & "</TD>" & vbCrLf
			ihtml = ihtml & "</TR>" & vbCrLf
		Case Else
		
	End Select
End Sub
'------------------------------------------------------------------------------------------------------------------------------------->	

'//SECTION FOUR//

'//COMPLETE THE HTML AND RETURN RESULTS//

'//Public Sub iFooter(pgselect)//
'//Public Sub pageinfo()//

'------------------------------------------------------------------------------------------------------------------------------------->	
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub iFooter(pgselect)	Public Sub iFooter(pgselect)	Public Sub iFooter(pgselect)	Public Sub iFooter(pgselect)	
	'	
	'	
	'************************************************************************************************************************************
Public Sub iFooter(pgselect)
If pgselect Then
	If Len(imsg) Then
	ihtml = ihtml & "<TR>" & vbCrLf
		ihtml = ihtml & "<TD	colspan=""4""><b>"
		ihtml = ihtml & imsg
		ihtml = ihtml & "</b></TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	End If
	ihtml = ihtml & "<TR bgcolor=""white"" STYLE=""height:3;"">" & vbCrLf
		ihtml = ihtml & "<TD	colspan=""4"">"
		ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	Call pageinfo()
ihtml = ihtml & "</TABLE>" & vbCrLf
End If			
	ihtml = ihtml & "</FORM>" & vbCrLf
	ihtml = ihtml & "</BODY>" & vbCrLf
	ihtml = ihtml & "</HTML>" & vbCrLf
End Sub																											
	'------------------------------------------------------------------------------------------------------------------------------------
	'	Public Sub pageinfo()	Public Sub pageinfo()	Public Sub pageinfo()	Public Sub pageinfo()	Public Sub pageinfo()		
	'	
	'	
	'************************************************************************************************************************************
Public Sub pageinfo()
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD	colspan=""4"">"
	ihtml = ihtml & "<B>Current Image Information for "&strpage&"</B>"
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD	align=""left"">"
	ihtml = ihtml & "Total Images:"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD colspan=""3"">"
	ihtml = ihtml & timg+1&"  [of "&pginfo(4)&"]"
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD align=""left"">"
	ihtml = ihtml & "Images Displayed:"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD colspan=""3"">"
	Call loadmod("displayed")
	ihtml = ihtml & displayed&" [of "&pginfo(0)&"]"
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD align=""left"">"
	ihtml = ihtml & "Auto Create Thumbnail:"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD colspan=""3"">"
	If pginfo(1) = 1 Then
		ihtml = ihtml & "Yes"
	Else
		ihtml = ihtml & "No"
	End If
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD	align=""left"">"
	ihtml = ihtml & "Thumnail Height:"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD	colspan=""3"">"
	ihtml = ihtml & pginfo(2)
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD	align=""left"">"
	ihtml = ihtml & "Thumbnail Width:"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD colspan=""3"">"
	ihtml = ihtml & pginfo(3)
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
ihtml = ihtml & "<TR>" & vbCrLf
	ihtml = ihtml & "<TD	align=""left"">"
	ihtml = ihtml & "&nbsp;"
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "<TD	colspan=""3"">"
	ihtml = ihtml & "&nbsp;"
	ihtml = ihtml & "</TD>" & vbCrLf
ihtml = ihtml & "</TR>" & vbCrLf
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------
'	End vbScript Source Code	End vbScript Source Code	End vbScript Source Code	End vbScript Source Code				    
'****************************************************************************************************************************************
%>