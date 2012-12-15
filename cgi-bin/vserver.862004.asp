<%
'------------------------------------------------------------------------------------------------------------------------>	
'<--------- STEP ONE // GRAB INFORMATION AND DESIGNATE THE ACTION
Public Sub ServerAction()
		   a = Request.QueryString("imagemgtup")
		   b = request.QueryString("boothumb")
		   icrop = request.QueryString("icrop")
		   x = Cstr(Request.Querystring("Submit1"))
		page = Cstr(Request.QueryString("hideme"))
   Call processinfo(x,page,a,b)
   Response.Redirect ("imagemgt.asp?strpage="&page)
End Sub
'------------------------------------------------------------------------------------------------------------------------>	
'<--------- STEP TWO // PROCESS INFORMATION BASED ON ACTION [MODIFY PLACEMENT|REMOVE IMAGE(S)] OR SEND TO APPROPRIATE ROUTINE [ADD IMAGE]
Public Sub processinfo(action,page,a,b)
		SELECT CASE action
			CASE "Modify Placement"
				a = Split(a,",")
				sql = "DELETE ImageOrder.* FROM Image_REL_Page "
				SQL = SQL&"INNER JOIN ImageOrder ON Image_REL_Page.REL_Image_PageID = ImageOrder.REL_Image_PageID "
				SQL = SQL&"WHERE Image_REL_Page.PageID='"& page &"';"
					Call oCon("oA",SQL,adOpenStatic,adLockPessimistic)
					sql = "SELECT * FROM IMAGEORDER;"
					Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
					For i = 1 to UBOUND(a)+1
						If isnumeric(a(i-1)) Then 
							rs.AddNew
							rs.fields("ImageOrder")=i
							rs.fields("REL_Image_PageID")=a(i-1)
							rs.update
						End If 
					NEXT	
					Call oCon("c","","","")
			Case "Remove Image(s)"												'//remove image information from the database and
																				'//(w/permissions) delete the image(s) from the server
					v = Split(a,",")											'//split the information out of the string
					ua=Ubound(v)
					ua = ua - 1	
					ua = ua/2													'//ubound of the image array [length of string-1 / 2]
					ReDim a(3,ua)
					For i = 0 to ua
						a(0,i)=v(2*i)											'//Image_REL_PageID
						a(1,i)=Trim(""&v((2*i)+1)&"")							'//Image Source For File Deletion ["SHARED"]
					Next
					For i = 0 To ua												'//For...Next the image information
					SQL = "SELECT Image_Inf.* FROM Image_Inf INNER JOIN Image_REL_Page ON Image_Inf.ImageID = Image_REL_Page.ImageID "
					SQL = SQL&"WHERE Image_REL_Page.REL_Image_PageID="&a(0,i)&";"
					Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)
						If Not RS.EOF Then						
							a(2,i) = rs("Mode")									'//Image Availability Mode ["EXCLUSIVE"|"SHARED"|"GLOBAL"]
							a(3,i) = rs("Thumb")								'//Has a Thumbnail Been Created?
						End If
					Call oCon("c","","","")
					Next
					For i = 0 To ua												'//For...Next the delete process for each image
					Call fsome(a(1,i),"pathonly",False)							'//set the file path variable
					Select Case a(2,i)
						Case "1"												'//exclusive (phase one) code...images cannot be 
																				'//used on other pages, so all images are deleted
																				'//START DATABASE SQL DELETION
																				'//DELETING image_inf.* image_rel_page.* HERE
					SQL = "DELETE Image_Inf.*, Image_REL_Page.* FROM Image_Inf INNER JOIN Image_REL_Page "
					SQL = SQL&"ON Image_Inf.ImageID = Image_REL_Page.ImageID "
					SQL = SQL&"WHERE Image_REL_Page.REL_Image_PageID="&a(0,i)&";"
					Call oCon("oA",SQL,adOpenStatic,adLockPessimistic)
																				'//DELETING image_order.*(by way of rs.delete HERE)
					SQL = "SELECT ImageOrder.* FROM ImageOrder WHERE ImageOrder.REL_Image_PageID="&a(0,i)&";"
					Call oCon("oA",SQL,adOpenStatic,adLockPessimistic)
						IF NOT RS.EOF THEN
							RS.Delete
							RS.Update
						END IF
					Call oCon("c","","","")
																				'//END DATABASE SQL Deletion
					Call fsome(filespec,"objectinfo",false)						'//create and use the filesystem object which handles the 
																				'//deletion of images in root web directory (SERVER)
					Call fsome(filespec,"logme",logmsg)							'//enables the log file creation
					filespec = ""												'//resets the file path
					strmsg = ""													'//resets the HTML client output
					logmsg = ""													'//resets the physical log
					CASE "2"
																				'shared mode code to come later
					CASE "3"
																				'global mode code to come later
					CASE ELSE
				End Select
				Next
			Case "Add Image"
				Call ASPUpload("add",page,a,b)									'//Calls Add Image functionality
				
			Case Else															'//Select Safety Case for Error Prevention
			
			End Select															
	End Sub	
'------------------------------------------------------------------------------------------------------------------------>	
'<--------- STEP TWO[b] // LOG REMOVE INFORMATION INTO TWO STRINGS [HTML|TXT] THEN PRINT TO PAGE AND FILE SYSTEM OBJECT TEXTFILE
Public Sub fsome(path,logyet,msg)
		Select Case logyet
			Case "pathonly"																		'CREATE THE PATH TO THE FILE
					fx = split(path,"://")
					fs = split(fx(1),"/")
					For f = 1 To Ubound(fs)
						filespec = filespec & "/" & fs(f)										'//create relative path to the file
					Next
					filespec = SERVER.MapPath(filespec)											'//get the physical drive/path to the file
			Case "objectinfo"																	'DELETE THE FILE(S) AND COMPILE INFO
					base = fso.getbasename(filespec)													
					if right(base,3)="_TH" then													'//if the filename is a thumbnail, remove
						base  = len(fso.getextensionname(filespec))+4							'//gets the length of the extension + _TH.
						base = len(filespec)-base
						ext  = fso.getextensionname(filespec)
						filespec = left(filespec,base)
						filespec = filespec & "." & ext
						response.Write (filespec & "<br><br>")
					end if
					ext = "."&fso.GetExtensionName(filespec)									'//file extention (i.e. .jpg)
					ffile = fso.GetFileName(filespec)											'//uses base+ext to compile filename
					f = Split(filespec,ffile)													'//gets the directory path to the file
					strmsg = strmsg&"SQL [DELETE] Relational Data:"&ffile&":"&now&"<br>"
					logmsg = logmsg&"SQL [DELETE] Relational Data:"&path&":"&now&"<br>"																														
								IF fso.FileExists(filespec) THEN								'//check for the image
									strmsg = strmsg&"File Exists:"&ffile&":"&now&"<br>"
									logmsg = logmsg&"FileExists:"&filespec&":"&now&"<br>"
									call fso.DeleteFile(filespec,true)							'//DELETE IMAGE FROM SERVER
									strmsg = strmsg&"Delete File:"&ffile&":"&now&"<br>"
									logmsg = logmsg&"DeleteFile:"&filespec&":"&now&"<br>"
								ELSE															'//If it is not on server
									strmsg = strmsg&"File Exists:FALSE:"&now&"<br>"
									logmsg = logmsg&"FileExists:FALSE:"&now&"<br>"
									strmsg = strmsg&"Delete File:FALSE:"&now&"<br>"
									logmsg = logmsg&"DeleteFile:FALSE:"&now&"<br>"						
								END IF
																										'//END IMAGE DELETION
																										'//BEGIN THUMBNAIL DELETION
								IF a(3,i) THEN															'//See if a thumbnail should exist
									strmsg = strmsg&"Thumbnailed:"&a(3,i)&":"&now&"<br>"		
									logmsg = logmsg&"Thumbnailed:"&a(3,i)&":"&now&"<br>"		
										base = fso.getbasename(filespec)
										ext  = fso.getextensionname(filespec)
										base = base&"_TH."&ext											'//Sets base variable to thumbnail
										filespec = fso.BuildPath(f(0),base)								'//Builds path to thumbnail
										IF fso.FileExists(filespec) THEN								'//Checks for physical thumbnail
											strmsg = strmsg&"Thumbnail Exists:"&base&":"&now&"<br>"	
											logmsg = logmsg&"ThumbnailExists:"&filespec&":"&now&"<br>"
											call fso.DeleteFile(filespec,true)							'//DELETES image from server
											strmsg = strmsg&"Delete Thumbnail:"&base&":"&now&"<br>"
											logmsg = logmsg&"DeleteThumbnail:"&filespec&":"&now&"<br>"	
										ELSE															'//If thumb is not on server
											strmsg = strmsg&"Thumbnail Exists:FALSE:"&now&"<br>"	
											logmsg = logmsg&"ThumbnailExists:FALSE:"&now&"<br>"
											strmsg = strmsg&"Delete Thumbnail:FALSE:"&now&"<br>"
											logmsg = logmsg&"DeleteThumbnail:FALSE:"&now&"<br>"	
										END IF
								END IF
																										'//END THUMBNAIL DELETION
							strmsg = strmsg&"Delete Process Complete:"&now&"<br>"
							logmsg = logmsg&"DeleteProcessComplete:"&now&"<br>"
							logmsg = logmsg&"WRITEBLANKLINES"&"<br>"
							response.Write(strmsg&"<br>")												'//Writes HTML client output
							
			CASE "logme"																				'//Logs File Deletion .txt file
					'server.ScriptTimeout=2700
					'	fd = fso.BuildPath(f(0),"fso")
					'	fdir = fd
					'SET ffo = fso.GetFolder(fdir)
					'CONST ForReading = 1, ForWriting = 2, ForAppending = 8
					'SET ftxt = fso.CreateTextFile(fdir&"RemoveLOG_"&DATE&"_"&session("sessionid")&".txt",true)
					'	ffol = ffo
					'	strmsg = Split(strmsg,"&<br>")
					'	FOR i = 0 to ubound(strmsg)
					'		IF NOT instr(strmsg(i),"WRITEBLANKLINES") THEN
					'			ftxt.WriteLine(strmsg(i))
					'		ELSE
					'			ftxt.WriteBlankLines(1)
					'		END IF 
					'	NEXT
					'ftxt.close()
					'server.ScriptTimeout=15
					'SET ffo = NOTHING
					'SET ftx = NOTHING
			Case Else
			
			End Select
End Sub
'------------------------------------------------------------------------------------------------------------------------>	
'<--------- STEP THREE // PROCESS THE ADD IMAGE INFORMATION AND UPLOAD FILE
Public Sub ASPUpload(operation,page,afilename,bboothumb)
	Dim sm,tmpdrv,nomore
	Dim iheight, iwidth
	Dim strAlt,localspec,objUpload,objsaveas,tmpfile,tmpspec,objfolder,tmpurl,tmpfolder,filename
	Dim sir,continue,boothumb														
	Call pginformation(page)																'//PUBLIC Sub Gather Page Information		//
	nomore = pginfo(8)																		'//Place max bytes value into variable		//
	filespec = afilename																	'//get filename from full path				//
	boothumb = bboothumb																	'//get thumbnailed boolean ["on"|"off"]		//
																							'//create File System Object for server		//
																							'//navigation and log creation				//
			base = fso.GetBaseName(filespec)												'//base file name(no extention)				//
			ext = "."&fso.GetExtensionName(filespec)										'//file extention (i.e. .jpg)				//
																							'//full filename [NOT NECESSARY]			//
	If Not IsNull (base) AND Not IsNull (ext) Then
		continue = true
	End If		
	If continue Then																		'//If is acceptable then continue upload	//
			filename = fso.GetFileName(filespec)											'//FileSystem Object [filename] remove above//
																							'//end filename extraction					//
	Select Case operation
		Case "add"
																							'//ASPSmartUpload START						//
			Set objUpload = Server.CreateObject("aspSmartUpload.SmartUpload")				'//Creates ASPSmartUpload Object			//
			strAlt = request.QueryString("stralt")											'//Grabs Alt tag value from url				//
			tmpspec = getroot("folder",pginfo(5))											'//Finds Root Folder from Request URL		//
			response.Write (tmpspec)
			tmpspec = fso.BuildPath(tmpspec,pginfo(5))										'//builds path [drive:]\[root][images path]	//
																							'//[images folder] path from root found in	//
																							'//default images folder field in database	//
			Set	tmpfolder = fso.GetFolder(tmpspec)											'//attaches the current folder to the		//
																							'//file system object						//
			objfolder = tmpfolder.path														'//physical path to the root server folder	//
			SQL = "SELECT * FROM Image_Inf WHERE Image_Inf.Image='" & filename & "';"		'//checks to see if an image with said name	//
																							'//already exists in database independant	//
																							'//of the page/user							//
																							'//ADD AS A SECONDARY CHECK					//
																							'//FSO CHECK OF FILENAME					//
			Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)								'//open connection/recordset				//
			If rs.eof Then																	'//if the file is unique continue			//
					rs.addnew																'//add new record							//
						rs.fields("Image") = TRIM(filename)									'//add the filename							//
						If boothumb = "on" Then												'//see if image should be thumbnailed		//
							rs.fields("Thumb") = 1											'//check the thumbnailed box in database	//
						End If										
						rs.fields("StrAlt") = StrAlt										'//enter the Alt text into the database		//
						rs.fields("Mode") = pginfo(7)										'//enter the Mode into the database			//
																							'//[EXCLUSIVE|SHARED|GLOBAL]				//
																							'//MODE IS FOR NOW ALWAYS EXCLUSIVE			//
																							'//OTHERS TO BE ADDED LATER					//
																							'//AT THAT TIME ADD OPTIONS FOR USER TO		//
																							'//MODIFY MODE[CURRENTLY IT ONLY CHECKS THE]//
																							'//[PAGE DEFAULTS FROM THE DATABASE]		//
					rs.update																'//Updates the recordset					//
					rs.Requery																'//requeries database to gather the newly	//
																							'//added record [CAN CHANGE THIS IF I WANT]	//
																							'//[IF THERE IS A NEED USE A DIFFERENT ADO]	//
																							'//[COMMAND|SHOULDN'T MATTER SINCE ONLY ONE]//
																							'//[ROW CAN EXIST]							//
					newid = rs.Fields("ImageID")											'//grab the new imageid						//
				Call oCon("c","","","")														'//close the recordset and connection		//
				SQL = "SELECT * FROM Image_REL_PAGE WHERE ImageID=" & newid & ";"			'//add the image-page relationship to the	//
																							'//database									//
				Call oCon("oA",SQL,adOpenStatic,adLockOptimistic)							'//open the connection						//
					If rs.EOF Then															'//if the relationship does not yet exist	//
						rs.AddNew															'//create the relationship					//
						rs.Fields("ImageID") = newid										'//add image								//
						rs.Fields("PageID") = page											'//add page									//
						rs.Update															'//update recordset							//
						continue = true														'//continue routine							//
					Else																	
						continue = false													'//if prior relationship exists QUIT [ERROR]//
																							'//this should not be possible unless		//
																							'//data corruption prior to or during		//
																							'//recordset updates						//
																							'//can use ADO to correct errors			//
																							'//NOT AN IMMEDIATE CONCERN...ADD LATER		//
					End If
				Call oCon("c","","","")												'//close connection and recordset			//
																			
			Else																			'//IF THE FILE IS NOT UNIQUE				//
				Call oCon("c","","","")												'//CLOSE CONNECTION AND RECORDSET			//
				continue = false															'//DO NOT CONTINUE THE ROUTINE				//
				imsg = imsg & "File Exists. Please Choose Another File or Rename.<br>"		'//CREATE ERROR STRING						//
				Response.Redirect ("imagemgt.asp?imsg="&imsg)								'//TRANSFER CONTROL BACK TO THE USER		//
			End If																			'//END CHECK FILE QUALIFICATIONS			//
					objUpload.AllowedFilesList = "jpg,jpeg,gif,png,bmp"						'//Allow only image uploads					//
					objUpload.DenyPhysicalPath = False						                '//Allow Virtual and Physical Server Path	//
					'objUpload.MaxFileSize = nomore											'//Set the Maximum File Size				//
					objUpload.Upload														'//UPLOAD the current image					//
					objsaveas = objUpload.Save(objfolder)									
				If Err Then																	'//Error Check								//
					imsg = imsg & "<B STYLE=""color:RED;"">Error: </B>" & Err.description	'//Compile Error imsg						//
				Else
					imsg = imsg & filename & " uploaded."									'//Compile Success imsg						//
				End If								
				If boothumb = "on" Then														'//If Create Thumbnail Box was checked		//
					Set sir = Server.CreateObject("SfImageResize.ImageResize")				'//Create SFImage Object					//
						sir.LoadFromFile(objfolder&"\"&filename)							'//Load the recently uploaded image			//
						filename = BASE & "_TH" & EXT										'//Create Thumbnail File Name				//
						iheight = sir.Height												'//Collects the image height				//
						iwidth = sir.Width													'//Collects the image width					//
						If iheight > iwidth Then											'//Check for the larger side and set it		//
							idocrop = "height"
							sir.Width = pginfo(2)	
																							'//equal to the largest thumbnail side		//				
						Elseif iwidth > iheight	Then										'//THIS IS TEMPORARY | BUT J&F DOES NOT		//
							idocrop = "width"
							sir.Height = pginfo(3)											'//REQUIRE THUMBNAIL CONFIGURATION			//				
						Elseif iwidth = iheight Then
							idocrop = false
							sir.Height = pginfo(3)
						End If		
						If EXT = "jpg" OR EXT = "jpeg" Then
						sir.Compression = 15
						End If
						sir.Algorithm = 2													'//AS OF NOW THE THUMB IS MAX SIDE = PIXELS //
						sir.DoResize														'//CREATE The Thumbnail						//
						sir.SaveToFile (objfolder & "\" & filename)							'//SAVE Thumbnail to the images dir			//
						If not idocrop = false Then
						sir.LoadFromFile (objfolder & "\" & filename)
						imsg = imsg & " icrop=" & icrop  & " idocrop=" & idocrop
						SELECT Case icrop													
							Case "topleft"
								sir.Crop 0,0,pginfo(2),pginfo(3)
							Case "center"
								if idocrop = "height" then
									tcrop = sir.Height/2
									bcrop = pginfo(3)/2
									tcrop = tcrop - bcrop
									sir.Crop 0,tcrop,pginfo(2),tcrop+pginfo(3)
								elseif idocrop = "width" then
									lcrop = sir.Width/2
									rcrop = pginfo(2)/2
									lcrop = lcrop - rcrop
									sir.Crop lcrop,0,lcrop+pginfo(2),pginfo(3)
								end if
							Case "botright"
								sir.Crop sir.Width-pginfo(2),sir.Height-pginfo(3),sir.Width,sir.Height
							Case else
								if icrop = "height" then
									tcrop = sir.Height/2
									bcrop = pginfo(3)/2
									tcrop = tcrop - bcrop
									sir.Crop 0,tcrop,pginfo(2),tcrop+pginfo(3)
								elseif icrop = "width" then
									lcrop = sir.Width/2
									rcrop = pginfo(2)/2
									lcrop = lcrop - rcrop
									sir.Crop lcrop,0,lcrop+pginfo(2),pginfo(3)
								end if
						END SELECT
						sir.SaveToFile (objfolder & "\" & filename)							'//SAVE Thumbnail to the images dir			//
						End If
					Set sir = Nothing														'//Set the object to nothing				//
		End If
			Set tmpfolder = Nothing															'//Set the folder object to nothing			//
'''		Case "update"																		'//SPACE FOR FUTURE UPDATE INFORMATION CASE //
'''			sql = "SELECT * FROM Image_Inf"													'//											//
'''			oCon("oA",SQL,adOpenStatic,adLockOptimistic)									'//											//
'''				rs.addnew																	'//											//
'''					rs.fields("Image") = cstr(thumbImage)									'//											//
'''					rs.fields("Thumb") = cstr(fullImage)									'//											//
'''					rs.fields("StrAlt") = CStr(productID)									'//											//
'''					rs.fields("Mode") = CStr(product)										'//											//
'''				rs.update																	'//											//
'''			rs.close																		'//											//
		Case Else
	End Select
		filespec = Null																		'//Set Public variable to null				//
			base = Null																		'//Set Public variable to null				//
			ext = Null																		'//Set Public variable to null				//
			filename = Null																	'//Set Public variable to null				//
			continue = False																'//Set Public variable to False				//
	End If 
End Sub
'------------------------------------------------------------------------------------------------------------------------>	
%>
