<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
 setimgwidth = 100
 setimgheight = 200
 
 apath=server.mappath("../Public/Data/")
 Set upl = Server.CreateObject("SoftArtisans.FileUp")
 upl.Path = apath

 If DataDecide = "Y" Then
  SQL = "Select UnitID from DataUnit Where ('"& date() &"' between BeginDate and EndDate) And DataType = N'"& DataType &"' And Language = N'"& Language &"' Order By ShowOrder"
  set rs = conn.execute(SQL)
	If not rs.EOF Then
	  NewOrderNo = 1
	  do while not rs.eof
	    NewOrderNo = NewOrderNo + 1
	    sql2="Update DataUnit set ShowOrder = "& NewOrderNo &" Where UnitID = " & rs("UnitID")
	    set rs2 = conn.execute(sql2)
	  rs.movenext
	  loop
	End if
 Else
  SQL = "Select UnitID from DataUnit Where DataType = N'"& DataType &"' And Language = N'"& Language &"' Order By ShowOrder"
  set rs = conn.execute(SQL)
	If not rs.EOF Then
	  NewOrderNo = 1
	  do while not rs.eof
	    NewOrderNo = NewOrderNo + 1
	    sql2="Update DataUnit set ShowOrder = "& NewOrderNo &" Where UnitID = " & rs("UnitID")
	    set rs2 = conn.execute(sql2)
	  rs.movenext
	  loop
	End if
 End IF
 
	ShowOrder = 1
	SQL = ""
	SQLForm = ""
	SQLVALUESForm = ""
	
	for each x in upl.form
		If upl.form(x) <> "" And left(x,3) = "xfn" Then
			SQLForm = SQLForm & " ," & mid(x,5)
			SQLVALUESForm = SQLVALUESForm & "," & NoHTMLCode(upl.form(x))
		ElseIf upl.form(x) <> "" And left(x,3) = "xin" Then
			SQLForm = SQLForm & " ," & mid(x,5)
			SQLVALUESForm = SQLVALUESForm & "," & upl.form(x)
		End IF
	next

	SQL = "INSERT INTO DataUnit (DataType, Language, EditDate, EditUserID, ShowOrder"& SQLForm &") VALUES "
	SQL = SQL & "(N'"& DataType &"',N'"& Language &"',N'"& Date() &"',N'"& session("UserID") & "'," & ShowOrder & SQLVALUESForm & ")" 
	set rs = conn.execute(SQL)

	SQL = "Select UnitID from DataUnit Where DataType = N'"& DataType &"' And Language = N'"& Language &"' And EditDate = N'"& Date() &"' And ShowOrder = 1"
	set rs = conn.execute(SQL)
		UnitID = rs("UnitID")

 for each xItem in upl.form
  if left(xItem,7)="Content" then
   fno=mid(xItem,8)
   ufile = upl.form("imagefile"& fno).UserFilename
   ofname = fname(ufile)
   Content = upl.Form("Content"& fno)
   ImageWay = upl.Form("ImageWay"& fno)
   Phase = upl.Form("No"& fno)

	if ofname<>"" then
	  if instr(ofname, ".")>0 then
	    fnext=mid(ofname, instr(ofname, "."))
	  else
	    fnext=""
	  end if
	  tstr = now()
	  nfname = year(tstr) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & fno & fnext
	  upl.form("imagefile"& fno).SaveAs nfname

'-----圖片大小調整開始
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
'-----圖片大小調整結束
    Else
     nfname = ""
     ofname = "" 
	End If

	SQL = "INSERT INTO DataContent (UnitID, Content, ImageFile, ImageWay, Position) "&_
		  "VALUES ("& UnitID &","& pkStr(Content) &",N'"& nfname &"',N'"& ImageWay &"',"& Phase &")"
    set rs = conn.Execute(SQL)

 end if
next
		
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
<script language=VBScript>
   alert "新增成功!"
<% If upl.Form("EMailCk") = "Y" Then %>
     document.location.href="BulletinEMail.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>"
<% Else %>             
     document.location.href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>"
<% End IF %>
</script>
