<%@ Language=VBScript %>
<%	option explicit %>
<% 
	'response.Buffer = true 
	'response.CacheControl = "Private" 
	response.ExpiresAbsolute = #1/8/1980#
	'response.Expires = -1
	server.ScriptTimeout = 20
%>
<!--#include file="adovbs.inc"-->
<%
'	VBScript Source Code
'---------------------------------------------------------------------------------------------------------------------------------------->
'<----------//OBJECT VARIABLES
	Dim CONN											'//ADODB Connection Object
	Private StrCon										'//STR	- Contains the ADODB Connection String	
	Private Root										'//STR	- Path From Root to cgi-bin
	Private ADB											'//STR	- Access Database
	Public FSO											'//Active Scripting File System Object
	Public RS											'//ADODB Recordset Object
		Set CONN = Server.CreateObject("Adodb.Connection")
		Set RS = Server.CreateObject("adodb.recordset")
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
'---------------------------------------------------------------------------------------------------------------------------------------->
'<----------//PUBLIC ARRAYS
	Public pginfo(8)									'//1Dimension - Holds the Page Parameter Information
	Public arimg()										'//1Dimension - Holds the ImageID for the RecordSet
	Public ocx()										'//[Line 599] - Used to Create the Body Onload
	Public arrpage										'//
	Public arr											'//
	Public arrel										'//3Dimension - Holds Image_REL Info * [NOT YET USED]
	Public v											'//
'---------------------------------------------------------------------------------------------------------------------------------------->
'<----------//COMMON VARIABLES
	Dim b,c,e,g,h,ihtml,column,SQL,colrs,text,text_array,text_count,text_counter,header,i,timg,ii
	Dim txtb,intb,plus,modinfo,subone,tmpurl,sm,tmpspec,x
	Dim y,a,d,irows,moveme
	Dim ua
	Root	=	"D:\PAGES\JFElectric\"
	ADB		=	"cgi-bin\mailinglist.mdb;"											'//Use For Access Database Connection	//
	With Conn
		'.ConnectionString = "DRIVER={Microsoft Access Driver (*.mdb)};" &_
		'					"DBQ=" & Server.MapPath ("/cgi-bin/mailinglist.mdb")
		.ConnectionString = "DSN=jfelectricsupply;"
		.Open
	End With
'---------------------------------------------------------------------------------------------------------------------------------------->
'//opens and closes connections to a database
	Public Sub ocon(action,SQL,b,c)
		Select Case action
			Case "oA"
				Call VGCloseConn("all")
			'	response.Write (sql & "," & b & "," & c)
				RS.Open SQL,Conn,b,c															'//If Closed Then Open Recordset		//

			Case "oS"																			'//Use For SQL Database Connection		//
		
			Case "c"																			'//Check That RecordSet is Open			//
				Rs.Close
			Case Else
			
		End Select
	End Sub
'---------------------------------------------------------------------------------------------------------------------------------------->
	Public Sub VGCloseConn(vgclose)
		SELECT CASE	vgclose
			CASE "all"
				If Not Rs.State = adStateClosed Then
					Rs.Close
				End If
			CASE Else
				If Rs.State = adStateOpen Then
					Rs.Close
				End If
			END SELECT
	End Sub
'---------------------------------------------------------------------------------------------------------------------------------------->
'<----------//[SERVER]SIMPLE RETURN PATH														'//[DRIVE:\][ROOT FOLDER]				//
	Public Sub pginformation(strpage)														
	SQL	= "SELECT *	from page_image_spec where page_image_spec.pageID='" & strpage & "';"
	Call oCon("oA",SQL,adOpenStatic,adLockReadOnly)
		If Not RS.EOF Then
			pginfo(0) =	RS("intmaximage")														'//Number of image displayed on strpage	//
			pginfo(1) =	RS("forcethumb")														'//Force the system To create thumbnails//
			pginfo(2) =	RS("intthumbWIDTH")														'//Thumbnail Width						//
			pginfo(3) =	RS("intthumbheight")													'//Thumbnail Height						//
			pginfo(4) = RS("inttotimage")														'//Total Image Uploads Per StrPage		//
			pginfo(5) = RS("DefaultFolder")														'//Default Image Folder					//
			pginfo(6) = RS("ForcePopup")														'//Force Thumbnail+Popup Web Interface	//
			pginfo(7) = RS("ForceMode")															'//Force the mode for the pages images	//
			pginfo(8) = RS("IntThumbSize")														'//Max Size [Bytes][i.e. 50000 = 50k]	//
		Else
		End If 
	Call ocon("c","","","")
	End Sub
'---------------------------------------------------------------------------------------------------------------------------------------->
'<----------//[SERVER]SIMPLE RETURN PATH														'//[DRIVE:\][ROOT FOLDER]				//
	Function getroot(what,whatpath)
		whatpath = Replace (whatpath,"/","\")
		tmpurl = request.ServerVariables("URL")	
		sm = server.MapPath(tmpurl)																'//Maps Server Path to current page		//
		tmpurl = Split(sm,"\")																	
		TmpSpec = Null																			'//Drive:,Folder,...,File.Extention		//
		For x = 0 To Ubound (tmpurl) - 1														'//Number of folders in the root path	//
			If Not isNull(TmpSpec) Then 
				TmpSpec = TmpSpec & tmpurl(x+1)
			Else
				TmpSpec = TmpSpec & tmpurl(x)&"\"&tmpurl(x+1)									'//Reconstructs [DRIVE:\][ROOT FOLDER]	//
			End If
			If FSO.FolderExists (TmpSpec) Then 
				If Right (TmpSpec,1) <> "\" Then
					TmpSpec = TmpSpec & "\"
				End If
				sm = TmpSpec & whatpath
				Select Case what
					Case "folder"
							If FSO.FolderExists (sm) Then
								Exit For
							End If
							If x = Ubound (tmpurl) - 1	Then									'//Check to see that folders exist		//
								imsg   = imsg & vbCrLf & "Error: The Folder Does Not Exist. Please Contact the System Administrator."
								Exit For
							End If
					Case "file"
							If FSO.FileExists (sm) Then 
								Exit For
							End If
							If x = Ubound (tmpurl) - 1	Then									'//Check to see that the file exists	//
								imsg   = imsg & vbCrLf & "Error: The File Does Not Exist. Please Contact the System Administrator."
								Exit For
							End If
					Case Else
								imsg   = imsg & vbCrLf & "Error: The Filetype is Unrecognized. Please Contact the System Administrator."
				End Select
				sm = Null
			End If
		Next
		getroot = TmpSpec
		tmpurl	= Null																			
		sm		= Null
		x		= Null																			'//CleanUP Variables					//
	End Function
	Function getthumb(src)
		ithumb = fso.GetBaseName(src)
		ithumb = ithumb & "_TH." & fso.GetExtensionName(src)
		getthumb = ithumb
	End Function
'---------------------------------------------------------------------------------------------------------------------------------------->
%>