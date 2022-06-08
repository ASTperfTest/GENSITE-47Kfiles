<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="椸ƺ@"
HTProgFunc="sW"
HTUploadPath=session("Public")+"Data/"
HTProgCode="GC1AP2"
HTProgPrefix="DsdXML" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%

Dim pKey
Dim RSreg
Dim formFunction
Dim Sql, SqlValue
Dim xNewIdentity
taskLable="sW" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

apath=server.mappath(HTUploadPath) & "\"
set xup = Server.CreateObject("UpDownExpress.FileUpload")
xup.Open 
function xUpForm(xvar)
on error resume next
	xStr = ""
	arrVal = xup.MultiVal(xvar)
	for i = 1 to Ubound(arrVal)
		xStr = xStr & arrVal(i) & ", "
'		Response.Write arrVal(i) & "<br>" & Chr(13)
	next 
	if xStr = "" then
		xStr = xup(xvar)
		xUpForm = xStr
	else
		xUpForm = left(xStr, len(xStr)-2)
	end if
end function

'function xUpForm(xvar)
'	xUpForm = request.form(xvar)
'end function

  	set htPageDom = session("codeXMLSpec")
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='3' or showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

if xUpForm("submitTask") = "ADD" then
	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else
	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if
%>
<% sub initForm() %>
<script language=vbs>
cvbCRLF = vbCRLF
cTabchar = chr(9)

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- sWɪw]ȩbo
<%
	for each xparam in allModel2.selectNodes("//fieldList/field[formList!='']") 
'		response.write "'" & nullText(xparam.selectSingleNode("fieldName")) & vbCRLF
		AddProcessInit xparam
	next
	if nullText(allModel2.selectSingleNode("//fieldList/field[fieldName='topCat']/showList"))<>"" then _
		response.write "reg.htx_topCat.value = """ & left(session("tDataCat"),1) & """" 
%>
end sub

sub window_onLoad
	clientInitForm
end sub

    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
sub keywordComp(xBodyNodeStr)
	if reg.htx_xKeyword.value = "" then
		reg.htx_xKeyword.value = xBodyNodeStr
	else	
		i=0
		redim compArray(1,0)
		xKeywordArray=split(reg.htx_xKeyword.value,",")
		xBodyArray=split(xBodyNodeStr,",")
		'----Nreg.htx_xKeyword.valuerPvsJXְ}C
		for i=0 to ubound(xKeywordArray)
		    xPos=instr(xKeywordArray(i),"*")
		    if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			xWeight=mid(xKeywordArray(i),xPos+1)
		    else
			xStr=trim(xKeywordArray(i))
			xWeight=""
		    end if	
		    redim preserve compArray(1,i)
		    compArray(0,i)=xStr
		    compArray(1,i)=xWeight
		next	
		'----NhysearchǦ^_r}CӻPreg.htx_xKeyword.value}Cv@,YsbhsJXְ}C
		for j=0 to ubound(xBodyArray)
		    insertFlag=true
		    xPos2=instr(xBodyArray(j),"*")
		    if xPos2<>0 then
			xStr2=Left(trim(xBodyArray(j)),xPos2-1)
			xWeight2=mid(xBodyArray(j),xPos2+1)
		    else
			xStr2=trim(xBodyArray(j))
			xWeight2=""
		    end if			    
		    for k=0 to ubound(compArray,2)		    
			if xStr2=compArray(0,k) then
		    	    compArray(1,k)=xWeight2			    
			    insertFlag=false
			    exit for
			end if
		    next
		    if insertFlag then
			redim preserve compArray(1,i)
		    	compArray(0,i)=xStr2
		    	compArray(1,i)=xWeight2
			i=i+1
		    end if
					
		next
		'----NXְ}Cꦨ̫r
		xKeywordStr=""
		for w=0 to ubound(compArray,2)
		    if compArray(1,w)<>"" then
		    	xKeywordStr=xKeywordStr+compArray(0,w)+"*"+compArray(1,w)+","
		    else
		    	xKeywordStr=xKeywordStr+compArray(0,w)+","
		    end if
		next	
		reg.htx_xKeyword.value = Left(xKeywordStr,Len(xKeywordStr)-1)	
	end if	
end sub
 
sub keywordMake()
	if reg.htx_xBody.value="" then exit sub
	xStr=replace(replace(reg.htx_xBody.value,VBCRLF,""),"&nbsp;","")
	xmlStr = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & cvbCRLF
	xmlStr = xmlStr & "<xBodyList>" & cvbCRLF
	xmlStr = xmlStr & ncTabchar(1) & "<xBody><![CDATA[" & xStr & "]]></xBody>" & cvbCRLF
	xmlStr = xmlStr & "</xBodyList>" & cvbCRLF
  	set oXmlReg = createObject("Microsoft.XMLDOM")
  	oXmlReg.async = false
  	oXmlReg.loadXML xmlStr
	if oXmlReg.parseError.reason <> "" then 
		alert "夣ŦXr榡!"
		exit sub
	end if  	
  	postURL = "<%=session("mySiteURL")%>/ws/ws_xBody.asp"  	
	set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
	set rXmlObj = CreateObject("MICROSOFT.XMLDOM")
	rXmlObj.async = false
	xmlHTTP.open "POST", postURL, false
	xmlHTTp.send oXmlReg.xml  
	rv = rXmlObj.load(xmlHTTP.responseXML)		
	if not rv then
		alert "rǦ^X{~!"
		exit sub
	end if
	set xBodyNode = rXmlObj.selectSingleNode("//xBody")
	if xBodyNode.text<>"" then keywordComp xBodyNode.text
end sub
   
sub htx_xKeyword_onChange()
	'----ˬdrݤ@ӦrHW
	xKeywordArray=split(reg.htx_xKeyword.value,",")
	for i=0 to ubound(xKeywordArray)
		'----hvŸ
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr=trim(xKeywordArray(i))
		end if
		if blen(xStr)<=2 then
			alert "C@rצܤ֤GӦr!"
			exit sub
		end if
	next	
	'----ˬd
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xURI = "<%=session("mySiteURL")%>/ws/ws_keyword.asp?xKeyword=" & reg.htx_xKeyword.value
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")
	if xKeywordNode.text<>"" then KeywordWTP xKeywordNode.text
end sub

sub KeywordWTP(xKeyworWTPdStr)
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xStr=""
	xReturnValue=""
	xKeywordWTPArray=split(xKeyworWTPdStr,";")
	for i=0 to ubound(xKeywordWTPArray)
            chky=msgbox("`NI"& vbcrlf & vbcrlf &"@[ "&xKeywordWTPArray(i)&" ]rsb, [JݼfM椤ܡH"& vbcrlf , 48+1, "Ъ`NII")
            if chky=vbok then        
		xURI = "<%=session("mySiteURL")%>/ws/ws_keywordWTP.asp?xKeyword=" & xKeywordWTPArray(i)
		oXML.load(xURI)		          		
       	    end If
	next
end sub       
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "аȥgu{0}vAoťաI"
	lMsg = "u{0}v׳̦h{1}I"
	dMsg = "u{0}v yyyy/mm/dd I"
	iMsg = "u{0}vƭȡI"
	pMsg = "u{0}v " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 䤤@ءI"
<%
	for each param in allModel2.selectNodes("//fieldList/field[formList!='' and inputType!='hidden']") 
		processValid param
	next
%>
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="DsdXMLForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;
	    <font size=2>isW--DD椸:<%=session("CtUnitName")%> / 椸:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>j</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">d</a>&nbsp;&nbsp;	    
	       <a href="<%=HTprogPrefix%>List.asp">^C</a>	    
	       <%end if%>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<% end sub '--- showHTMLTail() ------%>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	SQL = "Select * From Client Where ClientID = N'"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "uȤsv!!Эsإ߫Ȥs!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

function d6date(dt)     'ন~  999/99/99 ƫASmallDateTime ϥ
	if Len(dt)=0 or isnull(dt) then
	     d6date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     'ɹs
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d6date=xy &  xm &  xd
                  end if
end function

sub doUpdateDB()
	sql = "INSERT INTO  CuDTGeneric("
	sqlValue = ") VALUES(" 
	for each param in allModel.selectNodes("dsTable[tableName='CuDTGeneric']/fieldList/field[formList!='']") 
		if nullText(param.selectSingleNode("fieldName")) = "xImportant" _
				 AND xUpForm("xxCheckImportant")="Y" then
			sql = sql & "xImportant,"
			sqlValue = sqlValue & pkStr(d6date(date()),",")
		elseif nullText(param.selectSingleNode("identity"))<>"Y" then	
			processInsert param
		end if
	next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
'	response.write sql
'	response.end
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)
'	response.write sql & "<HR>"

	sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "(giCuItem,"
	sqlValue = ") VALUES(" & dfn(xNewIdentity)
	for each param in refModel.selectNodes("fieldList/field[formList!='']") 
		processInsert param
	next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
'response.write sql
'response.end
	conn.Execute(SQL)  
	'----rBz
	if xUpForm("htx_xKeyword")<>"" then
	    redim iArray(1,0)
	    xStr=""
	    xReturnValue=""
	    SQLInsert=""
	    xKeywordArray=split(xUpForm("htx_xKeyword"),",")
	    weightsum=0
	    for i=0 to ubound(xKeywordArray)
	    	redim preserve iArray(1,i)
		'----}rPvŸ
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			iArray(0,i)=xStr
			iArray(1,i)=mid(xKeywordArray(i),xPos+1)
		else
			xStr=trim(xKeywordArray(i))
			iArray(0,i)=xStr
			iArray(1,i)=1		
		end if	
		weightsum=weightsum+iArray(1,i)
	    next   
	    '----SQLr 
	    for k=0 to ubound(iArray,2)
	    	SQLInsert=SQLInsert+"Insert Into CuDTKeyword values("+dfn(xNewIdentity)+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW{</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("sWI")
<%
	nextURL = HTprogPrefix & "List.asp"
	if xUpForm("nextTask") = "CuCheckList.asp" then
%>
		window.open "<%=session("myWWWSiteURL")%>/ContentOnly.asp?CuItem=<%=xNewIdentity%>&ItemID=<%=session("ItemID")%>&userID=<%=session("UserID")%>"
<%		
	elseif xUpForm("nextTask") <> "" then
		nextURL = xUpForm("nextTask") & "?iCuItem=" & xNewIdentity
	end if
%>
	document.location.href = "<%=nextURL%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
