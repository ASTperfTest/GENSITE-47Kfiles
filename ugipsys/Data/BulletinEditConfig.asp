<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
 setimgwidth = 100
 setimgheight = 200

	UnitID = Request.QueryString("UnitID") 
	ContentID = Request.QueryString("ContentID")
	EditType = Request.QueryString("EditType")

	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	apath=server.mappath("../Public/Data/")
	upl.Path = apath

	If EditType = "Del" Then
		SQL = "Select ImageFile FROM DataContent Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
		 IF Not rs.EOF Then
		  do while not rs.eof
		  	If Not (rs("ImageFile") = "" or IsNull(rs("ImageFile"))) Then 
			 set dfile = server.createobject("scripting.filesystemobject") 
			 ifile = server.mappath("../Public/Data/" & rs("ImageFile")) 
			 dfile.deletefile(ifile)
			End If 
		  rs.movenext
		  loop
		 End If
		 
		 If DataDecide = "Y" Then
			SQL = "DELETE FROM DataContent Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
			SQL = "DELETE FROM DataUnit Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
		  SQL = "Select UnitID from DataUnit Where ('"& date() &"' between BeginDate and EndDate) And DataType = N'"& DataType &"' And Language = N'"& Language &"' Order By ShowOrder"
		  set rs = conn.execute(SQL)
			If not rs.EOF Then
			  NewOrderNo = 0
			  do while not rs.eof
			    NewOrderNo = NewOrderNo + 1
			    sql2="Update DataUnit set ShowOrder = "& NewOrderNo &" Where UnitID = " & rs("UnitID")
			    set rs2 = conn.execute(sql2)
			  rs.movenext
			  loop
			End if
		 ElseIf CatDecide = "Y" Then
			SQL = "Select CatID FROM DataUnit Where UnitID = "& UnitID
			set rs = conn.execute(SQL)
			 CatID = rs(0)
			SQL = "DELETE FROM DataContent Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
			SQL = "DELETE FROM DataUnit Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
		  SQL = "Select UnitID from DataUnit Where DataType = N'"& DataType &"' And Language = N'"& Language &"' Order By ShowOrder"
		  set rs = conn.execute(SQL)
			If not rs.EOF Then
			  NewOrderNo = 0
			  do while not rs.eof
			    NewOrderNo = NewOrderNo + 1
			    sql2="Update DataUnit set ShowOrder = "& NewOrderNo &" Where UnitID = " & rs("UnitID")
			    set rs2 = conn.execute(sql2)
			  rs.movenext
			  loop
			End if
		 Else
			SQL = "DELETE FROM DataContent Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
			SQL = "DELETE FROM DataUnit Where UnitID = "& UnitID 
			set rs = conn.execute(SQL)
		  SQL = "Select UnitID from DataUnit Where DataType = N'"& DataType &"' And Language = N'"& Language &"' Order By ShowOrder"
		  set rs = conn.execute(SQL)
			If not rs.EOF Then
			  NewOrderNo = 0
			  do while not rs.eof
			    NewOrderNo = NewOrderNo + 1
			    sql2="Update DataUnit set ShowOrder = "& NewOrderNo &" Where UnitID = " & rs("UnitID")
			    set rs2 = conn.execute(sql2)
			  rs.movenext
			  loop
			End if
		 End IF %>
		<script language="vbscript">
		 alert "本文刪除成功！"
		 document.location.href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>"
		</script>
<%	 Response.End 
	ElseIf EditType = "AddImg" Then
	  ImageWay = upl.form("ImageWay")
	  ufile = upl.form("imagefile").UserFilename
	  ofname = fname(ufile)
	  if instr(ofname, ".")>0 then
	    fnext=mid(ofname, instr(ofname, "."))
	  else
	    fnext=""
	  end if
	  tstr = now()
	  nfname = year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & fnext
	  upl.form("imagefile").SaveAs nfname
	  	Set ILIB = server.createobject("Overpower.Imagelib")
		picstr = apath &"\"& nfname
		ILIB.picturesize picstr,picw,pich
		 if lcase(fnext) = ".jpg" Then
		  fnoName = 3
		 Elseif lcase(fnext) = ".gif" Then
		  fnoName = 2
		 End IF 
		 IF (picw/setimgwidth) > (pich/setimgheight) and (picw/setimgwidth) > 1 Then
			ILIB.width = setimgwidth
			ILIB.height = int(setimgwidth*pich/picw)
			ILIB.InsertPicture picstr,0,0,true,setimgwidth,int(setimgwidth*pich/picw)
			ILIB.SavePicture picstr,fno,100,""
	 	 ElseIf (picw/setimgwidth) < (pich/setimgheight) and (pich/setimgheight) > 1 Then
			ILIB.width = int(setimgheight*picw/pich)
			ILIB.height = setimgheight
			ILIB.InsertPicture picstr,0,0,true,int(setimgheight*picw/pich),setimgheight
			ILIB.SavePicture picstr,fnoName,100,""
	 	 End IF
		SQL = "UPDATE DataContent SET ImageFile = N'"& nfname &"',ImageWay = N'"& ImageWay &"' Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = "新增圖片完成！"
		EditDataSave
	ElseIf EditType = "ChgImg" Then
	 SQL = "Select ImageFile FROM DataContent Where ContentID = "& ContentID
	 set rs = conn.execute(SQL)
	 set dfile = server.createobject("scripting.filesystemobject") 
	 ifile = server.mappath("../Public/Data/" & rs("ImageFile")) 
	 dfile.deletefile(ifile) 
	  ImageWay = upl.form("ChgImageWay")
	  ufile = upl.form("Chgimagefile").UserFilename
	  ofname = fname(ufile)
	  if instr(ofname, ".")>0 then
	    fnext=mid(ofname, instr(ofname, "."))
	  else
	    fnext=""
	  end if
	  tstr = now()
	  nfname = year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & fnext
	  upl.form("Chgimagefile").SaveAs nfname
	  	Set ILIB = server.createobject("Overpower.Imagelib")
		picstr = apath &"\"& nfname
		ILIB.picturesize picstr,picw,pich
		 if lcase(fnext) = ".jpg" Then
		  fnoName = 3
		 Elseif lcase(fnext) = ".gif" Then
		  fnoName = 2
		 End IF 
		 IF (picw/setimgwidth) > (pich/setimgheight) and (picw/setimgwidth) > 1 Then
			ILIB.width = setimgwidth
			ILIB.height = int(setimgwidth*pich/picw)
			ILIB.InsertPicture picstr,0,0,true,setimgwidth,int(setimgwidth*pich/picw)
			ILIB.SavePicture picstr,fno,100,""
	 	 ElseIf (picw/setimgwidth) < (pich/setimgheight) and (pich/setimgheight) > 1 Then
			ILIB.width = int(setimgheight*picw/pich)
			ILIB.height = setimgheight
			ILIB.InsertPicture picstr,0,0,true,int(setimgheight*picw/pich),setimgheight
			ILIB.SavePicture picstr,fnoName,100,""
	 	 End IF
		SQL = "UPDATE DataContent SET ImageFile = N'"& nfname &"',ImageWay = N'"& ImageWay &"' Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = "圖片變更完成！"
		EditDataSave
	ElseIf EditType = "ImgLeft" Then
		SQL = "UPDATE DataContent SET ImageWay ='Left' Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = "圖片位置變更完成！"
		EditDataSave
	ElseIf EditType = "ImgRight" Then
		SQL = "UPDATE DataContent SET ImageWay ='Right' Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = "圖片位置變更完成！"
		EditDataSave
	ElseIf EditType = "DelImg" Then
		SQL = "Select ImageFile FROM DataContent Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		 set dfile = server.createobject("scripting.filesystemobject") 
		 ifile = server.mappath("../Public/Data/" & rs("ImageFile")) 
		 dfile.deletefile(ifile) 
		SQL = "UPDATE DataContent SET ImageFile = Null Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = "圖片刪除成功！"
		EditDataSave
	ElseIf EditType = "DelStage" Then
		SQL = "DELETE FROM DataContent Where ContentID = "& ContentID 
		set rs = conn.execute(SQL)
		SQL = "Select Position,ContentID FROM DataContent Where UnitID = "& UnitID &" Order By Position"
		set rs = conn.execute(SQL)
		 IF Not rs.EOF Then
		  NewOrder = 0
		  do while not rs.eof
		   NewOrder = NewOrder + 1
			SQLCom = "UPDATE DataContent SET Position = "& NewOrder &" Where ContentID = "& rs("ContentID")
			set rs2 = conn.execute(SQLCom)
		  rs.movenext
		  loop
		 End If
		AlertMsg = "段落刪除成功！"
		EditDataSave
	ElseIf EditType = "AddUpStage" Then
	  ImageWay = upl.form("AddUsImageWay")
	  FormContent = upl.form("AddUsContent")
	  Phase = upl.form("AddUsnowps")
	  ufile = upl.form("AddUsimagefile").UserFilename
	  ofname = fname(ufile)
	  nfname = ""
	 if ofname<>"" then
	  if instr(ofname, ".")>0 then
	    fnext=mid(ofname, instr(ofname, "."))
	  else
	    fnext=""
	  end if
	  tstr = now()
	  nfname = year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & fnext
	  upl.form("AddUsimagefile").SaveAs nfname
	  	Set ILIB = server.createobject("Overpower.Imagelib")
		picstr = apath &"\"& nfname
		ILIB.picturesize picstr,picw,pich
		 if lcase(fnext) = ".jpg" Then
		  fnoName = 3
		 Elseif lcase(fnext) = ".gif" Then
		  fnoName = 2
		 End IF 
		 IF (picw/setimgwidth) > (pich/setimgheight) and (picw/setimgwidth) > 1 Then
			ILIB.width = setimgwidth
			ILIB.height = int(setimgwidth*pich/picw)
			ILIB.InsertPicture picstr,0,0,true,setimgwidth,int(setimgwidth*pich/picw)
			ILIB.SavePicture picstr,fno,100,""
	 	 ElseIf (picw/setimgwidth) < (pich/setimgheight) and (pich/setimgheight) > 1 Then
			ILIB.width = int(setimgheight*picw/pich)
			ILIB.height = setimgheight
			ILIB.InsertPicture picstr,0,0,true,int(setimgheight*picw/pich),setimgheight
			ILIB.SavePicture picstr,fnoName,100,""
	 	 End IF
		End If
		
		SQL = "Select Position,ContentID FROM DataContent Where UnitID = "& UnitID &" Order By Position"
		set rs = conn.execute(SQL)
		 IF Not rs.EOF Then
		  do while not rs.eof
		   If rs("Position") >= cint(Phase) Then
			SQLCom = "UPDATE DataContent SET Position = "& rs("Position") + 1 &" Where ContentID = "& rs("ContentID")
			set rs2 = conn.execute(SQLCom)
		   End If
		  rs.movenext
		  loop
		 End If
	 	 
		SQL = "INSERT INTO DataContent (UnitID, Content, ImageFile, ImageWay, Position) "&_
			  "VALUES ("& UnitID &","& pkStr(FormContent) &",N'"& nfname &"',N'"& ImageWay &"',"& Phase &")"
	    set rs = conn.Execute(SQL)
		AlertMsg = "新增段落完成！"
		EditDataSave
	ElseIf EditType = "AddEndStage" Then
	  ImageWay = upl.form("AddEsImageWay")
	  FormContent = upl.form("AddEsContent")
	  Phase = upl.form("AddEsnowps")
	  ufile = upl.form("AddEsimagefile").UserFilename
	  ofname = fname(ufile)
	  nfname = ""
	 if ofname<>"" then
	  if instr(ofname, ".")>0 then
	    fnext=mid(ofname, instr(ofname, "."))
	  else
	    fnext=""
	  end if
	  tstr = now()
	  nfname = year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & fnext
	  upl.form("AddEsimagefile").SaveAs nfname
	  	Set ILIB = server.createobject("Overpower.Imagelib")
		picstr = apath &"\"& nfname
		ILIB.picturesize picstr,picw,pich
		 if lcase(fnext) = ".jpg" Then
		  fnoName = 3
		 Elseif lcase(fnext) = ".gif" Then
		  fnoName = 2
		 End IF 
		 IF (picw/setimgwidth) > (pich/setimgheight) and (picw/setimgwidth) > 1 Then
			ILIB.width = setimgwidth
			ILIB.height = int(setimgwidth*pich/picw)
			ILIB.InsertPicture picstr,0,0,true,setimgwidth,int(setimgwidth*pich/picw)
			ILIB.SavePicture picstr,fno,100,""
	 	 ElseIf (picw/setimgwidth) < (pich/setimgheight) and (pich/setimgheight) > 1 Then
			ILIB.width = int(setimgheight*picw/pich)
			ILIB.height = setimgheight
			ILIB.InsertPicture picstr,0,0,true,int(setimgheight*picw/pich),setimgheight
			ILIB.SavePicture picstr,fnoName,100,""
	 	 End IF
		End If
		
		SQL = "Select Position,ContentID FROM DataContent Where UnitID = "& UnitID &" Order By Position"
		set rs = conn.execute(SQL)
		 IF Not rs.EOF Then
		  do while not rs.eof
		   If rs("Position") > cint(Phase) And rs("Position") <> cint(Phase) Then
			SQLCom = "UPDATE DataContent SET Position = "& rs("Position") + 1 &" Where ContentID = "& rs("ContentID")
			set rs2 = conn.execute(SQLCom)
		   End If
		  rs.movenext
		  loop
		 End If
	 	 
		SQL = "INSERT INTO DataContent (UnitID, Content, ImageFile, ImageWay, Position) "&_
			  "VALUES ("& UnitID &","& pkStr(FormContent) &",N'"& nfname &"',N'"& ImageWay &"',"& cint(Phase) + 1 &")"
	    set rs = conn.Execute(SQL)
		AlertMsg = "新增段落完成！"
		EditDataSave
	ElseIf EditType = "EditStage" Then
	    FormContent = upl.form("EsContent")
		SQL = "UPDATE DataContent SET Content = "& pkStr(FormContent) &" Where ContentID = "& ContentID
		set rs = conn.execute(SQL)
		AlertMsg = Content & "編修完成！"
		EditDataSave
	ElseIf EditType = "EditCat" Then
		SQL = "UPDATE DataUnit SET CatID = "& upl.Form("xin_CatID") &", EditUserID = N'"& session("UserID") &"', EditDate = N'"& Date() &"' Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
		AlertMsg = "類別編修完成！"
	ElseIf EditType = "EditData" Then
		SQL = "UPDATE DataUnit SET BeginDate = N'"& upl.Form("BeginDate") &"', EndDate = N'"& upl.Form("EndDate") &"', EditUserID =N'"& session("UserID") &"', EditDate = N'"& Date() &"' Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
		AlertMsg = "公佈時間編修完成！"
	ElseIf EditType = "EditSubject" Then
		SQL = "UPDATE DataUnit SET Subject = "& NoHTMLCode(upl.Form("Subject")) &", EditUserID = N'"& session("UserID") &"', EditDate = N'"& Date() &"' Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
		AlertMsg = Subject &"編修完成！"
	ElseIf EditType = "EditExtend_1" Then
		SQL = "UPDATE DataUnit SET Extend_1 = "& NoHTMLCode(upl.Form("Extend_1")) &", EditUserID = N'"& session("UserID") &"', EditDate = N'"& Date() &"' Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
		AlertMsg = Extend_1 &"編修完成！"
	End IF
	
	
	Function EditDataSave()
	 	SQL = "UPDATE DataUnit SET EditUserID = N'"& session("UserID") &"', EditDate = N'"& Date() &"' Where UnitID = "& UnitID
		set rs = conn.execute(SQL)
	End Function

Function NoHTMLCode(DataCode)
  NewData = "" 
  If DataCode <> "" Then
   NewDate = Replace(DataCode,"'","''")
   NewDate = Replace(NewDate,"<","&lt;")
   NewDate = Replace(NewDate,">","&gt;")
   NoHTMLCode = "'" & NewDate & "'"
  End If
End Function

Function fname(fstr)
  tmpstr=trim(fstr)
  cpos=instrrev(tmpstr, "\")
  if cpos=0 then
    fname=tmpstr
  else
    fname=mid(tmpstr,cpos+1)
  end if    
End Function

FUNCTION pkStr(s)     
  if trim(s)="" then     
    pkStr="NULL"     
  else     
    s=trim(s)     
    pos = InStr(s, "'")     
    While pos > 0     
      s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)     
      pos = InStr(pos + 2, s, "'")     
    Wend     
    pkStr="'" & s & "'"     
  end if     
END FUNCTION     	
%>
	
	<script language="vbscript">
	 alert "<%=AlertMsg%>"
	 document.location.href="BulletinEdit.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>"
	</script>