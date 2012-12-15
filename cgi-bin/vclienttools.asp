<%
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
End Sub
Function loadpic(page)
	Dim rget,looper,intprei,ifolder,iname
	ifolder = pginfo(5) & "/"
		SQL="SELECT ImageOrder.REL_Image_PageID, ImageOrder.ImageOrder, Image_Inf.* FROM (Image_Inf INNER JOIN Image_REL_Page ON "
		SQL=SQL&"Image_Inf.ImageID = Image_REL_Page.ImageID) INNER JOIN ImageOrder ON Image_REL_Page.REL_Image_PageID = "
		SQL=SQL&"ImageOrder.REL_Image_PageID WHERE Image_REL_Page.PageID='"&page&"' ORDER BY ImageOrder;"
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
								iname = rget(3,i)
								iname = fso.getbasename(iname) & "_TH." & fso.getextensionname(iname)
								iname = ifolder & iname
								%>intimgorder(<%=ii+1%>,'<%=iname%>',<%=pginfo(0)%>,'<%=rget(3,i)%>',<%=rget(2,i)%>,<%=rget(0,i)%>,<%=rget(1,i)%>,'<%=rget(5,i)%>');<%
							End If
							Next
						Next
						Response.Write ("""")
					End If 
			Else
			End If
End Function

Function showpics(page)
	
	For i = 1 to pginfo(0)
	ihtml =			"<TR> " & vbCrLf
	ihtml = ihtml & "<TD height="""&pginfo(3)*1.15&"""> " & vbCrLf
	Response.Write (ihtml)
	ihtml = ""
	%><IMG name="" border="0" onclick="vgpop(document.all['rimage'+<%=i%>].name,'<%=pginfo(2)%>','<%=pginfo(3)%>','<%=page%>');" STYLE="border-color:white;cursor:pointer;" height="0" ID="rimage<%=i%>"><%
	ihtml = ihtml & "</TD>" & vbCrLf
	ihtml = ihtml & "</TR>" & vbCrLf
	Response.Write (ihtml)
	ihtml = ""
	Next
End Function
%>