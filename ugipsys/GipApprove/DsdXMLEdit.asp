<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="ƼfZ"
HTProgFunc="s"
HTUploadPath=session("Public")+"Data/"
HTProgCode="GC1AP2"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="s" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

apath=server.mappath(HTUploadPath) & "\"
'response.write apath
'response.end
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

if request.querystring("S")="Approve" then	'----Ssession("codeXMLSpec")
	sql = "SELECT u.*,b.sBaseTableName FROM CuDTGeneric AS n LEFT JOIN CtUnit AS u ON u.CtUnitID=n.iCTUnit" _
		& " Left Join BaseDSD As b ON n.iBaseDSD=b.iBaseDSD" _
		& " WHERE n.iCuItem=" & request.querystring("iCuItem")
	set RS = Conn.execute(sql)
	session("CtUnitID") = RS("CtUnitID")
	session("ctUnitName") = RS("CtUnitName")
	session("iBaseDSD") = RS("iBaseDSD")
	session("fCtUnitOnly") = RS("fCtUnitOnly")
	if isNull(RS("sBaseTableName")) then
		session("sBaseTableName") = "CuDTx" & session("iBaseDSD")
	else
		session("sBaseTableName") = RS("sBaseTableName")
	end if	
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    	'----XCtUnitX???? xmlSpecɮ(Y䤣hdefault), èfieldSeqƧǦsJsession
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml") 	
    	if fso.FileExists(filePath) then
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml")
    	else
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
    	end if 
'	response.write LoadXML & "<HR>"
'	response.end
	xv = htPageDom.load(LoadXML)
'	response.write xv & "<HR>"
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if 	
    	set root = htPageDom.selectSingleNode("DataSchemaDef")
    	'----Load XSL˪O
    	set oxsl = server.createObject("microsoft.XMLDOM")
   	oxsl.async = false
   	xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    		
    	'----ƻsSlavedsTable,è̶ഫ
	set DSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement    
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&session("sBaseTableName")&"']")
    	'----ƻsCuDTGenericdsTable,è̶ഫ
    	set GenericNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDTGeneric']")       	


  	set session("codeXMLSpec") = htPageDom
  	'----VXfield
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(htPageDom.transformNode(oxsl))
	set session("codeXMLSpec2") = nxml0	
'response.write "Hello"
'response.end
end if

  	set htPageDom = session("codeXMLSpec")
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='3' or showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next	  	  	
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("Ƨs\I")
	end if

elseif xUpForm("submitTask") = "DELETE" then
	sql = "DELETE FROM CuDTGeneric WHERE iCuitem=" & pkStr(request.querystring("iCuItem"),"")
	conn.execute sql

	sql = "DELETE FROM " & nullText(refModel.selectSingleNode("tableName")) _
		& " WHERE giCuitem=" & pkStr(request.querystring("iCuItem"),"")
	conn.Execute SQL
	showDoneBox("ƧR\I")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.*, ghtx.* FROM " & nullText(refModel.selectSingleNode("tableName")) _
		& " AS htx JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCUItem "_
		& " WHERE ghtx.iCuItem=" & pkStr(request.queryString("iCuItem"),"")
	xFrom = nullText(refModel.selectSingleNode("tableName")) 
	Set RSreg = Conn.execute(sqlcom)
	pKey = "iCuItem=" & RSreg("iCuItem")


	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
'	response.write sqlcom
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		xqqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		qqRS = replace(xqqRS,chr(10),chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<%Sub initForm() %>
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

     sub clientInitForm
<%
	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
		EditProcessInit param
	next
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

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "<%=HTUploadPath%>" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addLogo(xname)	'sWlogo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'sWlogo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub

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
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
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
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("`NI"& vbcrlf & vbcrlf &"@ATwRƶܡH"& vbcrlf , 48+1, "Ъ`NII")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="DsdXMLFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>sת</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td class="FormName"><%=HTProgCap%>&nbsp;
	    <font size=2>i<%=HTProgFunc%>j</td>
    	<td class="FormLink" valign="top" align=right>
			<A href="Javascript: window.history.back();" title="^e">^e</A>
    	</td>
      	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>        
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = N'"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "uΤ@sv!!ЭsJȤΤ@s!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

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

Sub doUpdateDB()
	xn = 0
	sql = "UPDATE " & nullText(refModel.selectSingleNode("tableName")) & " SET "
	for each param in refModel.selectNodes("fieldList/field[formList!='']") 
		processUpdate param
		xn = xn + 1
	next
	sql = left(sql,len(sql)-1) & " WHERE giCuItem=" & pkStr(request.queryString("iCuItem"),"")
'	response.write sql & "<HR>"
'	response.end

	if xn>0 then	conn.Execute(SQL)  
	
	sql = "UPDATE CuDTGeneric SET "
	for each param in allModel.selectNodes("dsTable[tableName='CuDTGeneric']/fieldList/field[formList!='' and identity!='Y']") 
		if nullText(param.selectSingleNode("fieldName")) = "xImportant" _
				 AND xUpForm("xxCheckImportant")="Y" then
			sql = sql & "xImportant=" & pkStr(d6date(date()),",")
		else
			processUpdate param
		end if
	next
	sql = left(sql,len(sql)-1) & " WHERE iCuItem=" & pkStr(request.queryString("iCuItem"),"")
'	response.write sql & "<HR>"
'	response.end
	conn.Execute(SQL)  
	'----rBz
	'----R
	SQLDelete="Delete CuDTKeyword where iCuItem=" & pkStr(request.queryString("iCuItem"),"")
	conn.execute(SQLDelete)
	'----AsW
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
	    	SQLInsert=SQLInsert+"Insert Into CuDTKeyword values("+dfn(request.queryString("iCuItem"))+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if		
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>sת</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
