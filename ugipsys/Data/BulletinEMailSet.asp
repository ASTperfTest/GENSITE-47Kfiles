<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
UnitID = request.querystring("UnitID")

function chbr(s)
  if isNull(s) OR s="" then
    chbr=""
  else
    chbr=replace(s, vbcrlf, "<br>")
  end if
end function

function qqs(s)
  if isNull(s) OR s="" then
    qqs=""
  else
    xqqs=replace(s,chr(34),chr(34) & "&chr(34)&" &chr(34))
    qqs = replace(xqqs,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
  end if
end function

 SQLTitel = "SELECT E_ComName, C_ComName, EMail FROM COMPANY"
 Set RSTitle = conn.execute(SQLTitel)
  If Not RSTitle.EOF Then
   if not (ISNULL(RSTitle("C_ComName")) or RSTitle("C_ComName")="") then
   	E_ComName = RSTitle("C_ComName")   	
   else
   	E_ComName = RSTitle("E_ComName")
   end if
    EMail = RSTitle("EMail")
  Else
   E_ComName = session("corpCName")
   EMail = ""
  End IF    

	SQLCom = "SELECT InforUser.EnLName, InforUser.EnFName, DataUnit.BeginDate, "&_
		     "DataUnit.Subject "&_
			 "FROM DataUnit left JOIN "&_
		     "InforUser ON DataUnit.EditUserID = InforUser.UserID Where UnitID = "& UnitID
	SET RS = conn.execute(SQLCom)

	SQL="SELECT * FROM DataContent Where UnitID = "& UnitID &" Order By Position"
	SET RS2=conn.execute(sql)

Subject= rs("Subject")
HTMLSrc = "<html><head><link rel=stylesheet href=http://wcnin.neterprise.com.tw/PageStyle/system1.css></head><body>"
HTMLSrc = HTMLSrc & "<center><table border=0 width=500 cellspacing=0 cellpadding=0 class=c12-1>"
HTMLSrc = HTMLSrc & "<tr><td class=c12-3>"& rs("Subject") &"</td></tr>"

 if not RS2.EOF then
   while not RS2.EOF
	HTMLSrc = HTMLSrc & "<tr><td>"
    xPath = "http://wcnin.neterprise.com.tw" & session("Public")
    if not (ISNULL(RS2("ImageFile")) or RS2("ImageFile")="") then HTMLSrc = HTMLSrc & "<img src="& xPath & RS2("ImageFile") &" border=0 align="& rs2("ImageWay") &">"
   	HTMLSrc = HTMLSrc & chbr(RS2("Contents")) & "</td></tr>"
   RS2.movenext
   wend
 End if 
 
HTMLSrc = HTMLSrc &"<tr><td><br><br><br><br>" & E_ComName &"<br><a href=http://"& session("NPWebURL") &">http://"& session("NPWebURL") & "</a></td></tr>"
HTMLSrc = HTMLSrc & "</table></center>" 

 E_ComName = "=?utf-8?B?" & eb64Str(E_ComName) & "?="
 NewEmail = DelHTMLCode(EMail)

 If IsNull(NewEmail) = True or instr(NewEmail,"@") = 0 Then
  mailacc = E_ComName
 Else
  mailacc = E_ComName & "<" & NewEmail & ">"
 End If

 
  ClientEMailAddr = multi(Request.Form("MailCk"))	
   for x = 1 to ubound(ClientEMailAddr) 
       Set objNewMail = CreateObject("CDONTS.NewMail") 
       objNewMail.mailFormat = 0
       objNewMail.BodyFormat = 0 
       call objNewMail.Send(mailacc, ClientEMailAddr(x), Subject, HTMLSrc )
       Set objNewMail = Nothing
   Next
   
 Function DelHTMLCode(CodeData)
  NewData = CodeData
  While Instr(NewData,"<") > 0 And Instr(NewData,">") > 0
    TempText = Mid(NewData,Instr(NewData,"<"))
    DelCode = Mid(NewData,Instr(NewData,"<"),Instr(TempText,">"))
    Sno = Instr(NewData,"<") + 1 
    While Instr(2,DelCode,"<") > 0 
     TempText1 = Instr(Sno,NewData,"<")
     TempText = Mid(NewData,TempText1)
     DelCode = Mid(NewData,TempText1,Instr(TempText,">"))
     Sno = Sno + 1
    Wend
    NewData = Replace(NewData, DelCode,"")
  Wend
  DelHTMLCode = NewData
 End Function

function multi(tstr)
  dim outstr()
  to_no=1
  while len(tstr) > 0
    pos=instr(tstr, ",")
    if pos = 0 then
      redim preserve outstr(to_no)
      outstr(to_no) = tstr
      tstr = ""
    else
      redim preserve outstr(to_no)
      outstr(to_no) = left(tstr, pos-1)
      tstr=trim(mid(tstr, pos+1))
      to_no=to_no+1
    end if
  wend
  multi = outstr
end function

'====>> base 64 encoding

function eb64Str(bs)
	dim bv(4)
	dim cbytes(100)

	bslen = len(bs)
	cbuf = bs
	crStr = ""
	caCount = 0

	for i = 1 to bslen
		xasc = asc(mid(bs,i,1))
		if xasc < 0 then
			xasc = xasc + 65536
			cBytes(caCount) = int(xasc / 256)
			cBytes(caCount+1) = xasc mod 256
			caCount = caCount + 2
		else
			cBytes(caCount) = xasc
			caCount = caCount + 1
		end if
	next
	xcaCount = caCount
	while xcaCount mod 3 <> 0
		cBytes(xcaCount) = 0
		xcaCount = xcaCount +1
	wend
	
	crStr = ""
	c3 = int((caCount) / 3)
	for x3i = 0 to c3-1
		p = x3i*3
		crStr = crStr & e64(int(cBytes(p)/4))
		crStr = crStr & e64((cBytes(p) mod 4) * 16 + int(cBytes(p+1)/16))
		crStr = crStr & e64((cBytes(p+1) mod 16) * 4 + int(cBytes(p+2)/64))
		crStr = crStr & e64((cBytes(p+2) mod 64))
	next
	p = p+3
	while p < caCount
		if p mod 3 = 0 then
			crStr = crStr & e64(int(cBytes(p)/4))
			crStr = crStr & e64((cBytes(p) mod 4) * 16 + int(cBytes(p+1)/16))
		elseif p mod 3 = 1 then
			crStr = crStr & e64((cBytes(p) mod 16) * 4 + int(cBytes(p+1)/64))
		end if
		p = p+1
	wend
	while p mod 3 <> 0
		crStr = crStr & "="
		p = p+1
	wend
	
	eb64Str = crStr
end function

function e64(n)
	if n < 26 then
		e64 = chr(n+65)
	elseif n < 52 then
		e64 = chr(n+71)
	elseif n < 62 then
		e64 = chr(n-4)
	elseif n = 62 then
		e64 = chr(43)
	elseif n = 63 then
		e64 = chr(47)
	else
		e64 = "="
	end if
end function

'====>> base 64 encoding End
%>
<script language=VBScript>
    chky=msgbox("E-Mail 發送成功!"& vbcrlf & vbcrlf &"　繼續發信？"& vbcrlf , 48+1, "請注意！！")
    if chky=vbok then             
     document.location.href="BulletinEMail.asp?Language=<%=Language%>&DataType=<%=DataType%>&UnitID=<%=UnitID%>"
    Else               
	 document.location.href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>"
    end if  
</script>

