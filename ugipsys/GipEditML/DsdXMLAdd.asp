<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="新增"
HTUploadPath=session("Public")+"data/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<%
Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)

	Set objNewMail = CreateObject("CDONTS.NewMail") 
	objNewMail.MailFormat = 0
	objNewMail.BodyFormat = 0 
	call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)

	Set objNewMail = Nothing
End Function
function FilmRelated(xfunc,xTable,xType,xiCuItem,xFieldStr)
	Set Conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn2.Open session("ODBCDSN")
'Set Conn2 = Server.CreateObject("HyWebDB3.dbExecute")
Conn2.ConnectionString = session("ODBCDSN")
Conn2.ConnectionTimeout=0
Conn2.CursorLocation = 3
Conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

    	xxiCuItem=xiCuItem
    	if xTable="Corp" then 		'----Corp處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmCorpInfo where FilmNo="&xxiCuItem&" and CompanyType=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
	        xStr=trim(xKeywordArray(i))
		SQL="Select gicuitem from CorpInformation AI Left Join CuDtGeneric CDT " & _
			" ON AI.gicuitem=CDT.iCuITem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmCorpInfo values("&xxiCuItem&","&RSC(0)&",'"&xType&"',null,null,null);"
		end if
	    next     
    	elseif xTable="Actor" then 	'----People處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmPeopleInfo where FilmNo="&xxiCuItem&" and RoleInfo=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
		'----取最後括號
		xPos=instrRev(xKeywordArray(i),"(")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			xStrPar="'"+mid(trim(xKeywordArray(i)),xPos)+"'"
		else
			xStr=trim(xKeywordArray(i))
			xStrPAr="null"
		end if
		SQL="Select gicuitem from ActorInformation AI Left Join CuDtGeneric CDT " & _
			" ON AI.gicuitem=CDT.iCuITem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmPeopleInfo values("&xxiCuItem&","&RSC(0)&",'"&xType&"',"&xStrPAr&",null,null,null);"
		end if
	    next 
    	end if	
	if SQLI<>"" then conn2.execute(SQLI)  
'    	conn2.close
    	FilmRelated = ""
end function

Dim hyftdGIPStr
Dim pKey
Dim RSreg
Dim formFunction
Dim Sql, SqlValue
Dim xNewIdentity
Dim xMMOFolderID,MMOPath,xFTPIPMMO,xFTPPortMMO,xFTPIDMMO,xFTPPWDMMO,MMOFTPfilePath

taskLable="新增" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

'----ftp參數處理
	FTPErrorMSG=""
	FTPfilePath="public/data"
	SQLP = "Select * from UpLoadSite where upLoadSiteId='file'"
	Set RSP = conn.execute(SQLP)
	if not RSP.EOF  then
   		xFTPIP = RSP("UpLoadSiteFTPIP")
   		xFTPPort = RSP("UpLoadSiteFTPPort")
   		xFTPID = RSP("UpLoadSiteFTPID")
   		xFTPPWD = RSP("UpLoadSiteFTPPWD")
   	end if
'----ftp參數處理end 
	
apath=server.mappath(HTUploadPath) & "\"

if request.querystring("phase")<>"add" then
Set xup = Server.CreateObject("TABS.Upload")
xup.codepage=65001
xup.Start apath
else
Set xup = Server.CreateObject("TABS.Upload")
end if
function xUpForm(xvar)
xUpForm = xup.form(xvar)
end function

set htPageDom = session("codeXMLSpec")
set refModel = htPageDom.selectSingleNode("//dsTable")
set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
'----940215Film關聯處理session
session("FilmRelated_CorpActor")=nullText(allModel.selectSingleNode("FilmRelated_CorpActor"))
session("Country_related")=nullText(allModel.selectSingleNode("Country_related"))

if request.querystring("showType")="" then	'----一般資料
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xShowTypeStr="一般資料式"
  	xShowType="1"
elseif request.querystring("showType")="2" then	'----URL連結
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xShowTypeStr="URL連結式"
  	xShowType="2"
elseif request.querystring("showType")="3" then	'----檔案下載
	set allModel2 = session("codeXMLSpec2").documentElement   
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='4']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xShowTypeStr="檔案下載式"
  	xShowType="3"
elseif request.querystring("showType")="4" or request.querystring("showType")="5" then	'----引用資料
	'----被引用資料區	
	SQLG="Select CDG.iBaseDSD,CDG.iCTUnit,B.sBaseTableName from CuDtGeneric CDG Left Join BaseDSD B on CDG.iBaseDSD=B.iBaseDSD " & _
		"where iCUItem=" & request.querystring("iCuItem")
	set RSC=conn.execute(SQLG)
	if isNull(RSC("sBaseTableName")) then
		xBaseTableName = "CuDTx" & RSC("iBaseDSD")
	else
		xBaseTableName = RSC("sBaseTableName")
	end if		
	sqlCom = "SELECT htx.*, ghtx.*, xrefNFileName.oldFileName AS fxr_fileDownLoad FROM " & xBaseTableName _
		& " AS htx JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.gicuitem "_
		& " LEFT JOIN imageFile AS xrefNFileName ON xrefNFileName.newFileName = ghtx.fileDownLoad " _
		& " WHERE ghtx.iCUItem=" & pkStr(request.queryString("iCuItem"),"")
	Set RSreg = Conn.execute(sqlcom)	
	set htPageDomRef = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomRef.async = false
	htPageDomRef.setProperty("ServerHTTPRequest") = true		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default)
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(RSC("iCTUnit")) & ".xml") 	
	
    	if fso.FileExists(filePath) then
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(RSC("iCTUnit")) & ".xml")
    	else
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & RSC("iBaseDSD") & ".xml")
    	end if 
	'response.write LoadXML & "<HR>"
	'response.end
	xv = htPageDomRef.load(LoadXML)	
	set allModelRef = htPageDomRef.selectSingleNode("//DataSchemaDef")	
'response.write request.querystring("iCuItem")
'response.end
	'----新增資料區
    	'----Load XSL樣板
    	set oxsl2 = server.createObject("microsoft.XMLDOM")
   	oxsl2.async = false
   	xv = oxsl2.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    	
	set nxml2 = server.createObject("microsoft.XMLDOM")
	nxml2.LoadXML(htPageDom.transformNode(oxsl2))
	set allModel2 = nxml2.documentElement  
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='3']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xShowTypeStr="引用資料式"
  	xShowType=request.querystring("showType")
end if
'response.write "<XMP>"+allModel2.xml+"</XMP>"
'response.end	
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
'		xp = instr(xqqRS,vbCRLF&vbCRLF)
'		while xp > 0
'			xqqRS = left(xqqRS,xp-1) & mid(xqqRS,xp+4)
'			xp = instr(xqqRS,vbCRLF&vbCRLF)
'		wend
		xqqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		xqqRS = replace(xqqRS,chr(13),"")
		qqRS = replace(xqqRS,chr(10),"")
	end if
end function
%>
<% sub initForm() %>
<script language=vbs>
cvbCRLF = vbCRLF
cTabchar = chr(9)


Dim CanTarget
Dim followCanTarget

<% if session("documentDomain") <> "" then %>                        
 document.domain = "<%=session("documentDomain")%>"
<% end if %>

<%
	for each xCode in allModel.selectNodes("scriptCode")
		response.write replace(replace(xCode.text,chr(10),chr(13)&chr(10)),"zzzzzzmySiteURL",session("mySiteMMOURL"))
	next
%>

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub

sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.showType.value="<%=xShowType%>"
<%
	for each xparam in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    	if not (nullText(xparam.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") then
		    AddProcessInit xparam
	    	end if
	next	
	if request.querystring("showType")="4" then%>
	    reg.refID.value="<%=request.querystring("iCuItem")%>"
<%
	    for each param in allModel2.selectNodes(("//fieldList/field[formList!='' and fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit' and fieldName!='idept' and fieldName!='fctupublic' and fieldName!='ieditor' and fieldName!='deditDate' and fieldName!='createdDate' and fieldName!='showType' and fieldName!='refId']"))
	      	if nullText(allModelRef.selectSingleNode("//fieldList/field[fieldName='"&param.selectSingleNode("fieldName").text&"']"))<>"" then
			EditProcessInit param
		end if
	    next
	end if

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
		'----將reg.htx_xKeyword.value拆成字串與權重存入合併陣列中
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
		'----將hysearch傳回之內文斷詞字串陣列拿來與reg.htx_xKeyword.value陣列逐一比較,若不存在則存入合併陣列中
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
		'----將合併陣列串成最後字串
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
		alert "內文不符合字串比較格式!"
		exit sub
	end if  	
  	postURL = "<%=session("mySiteMMOURL")%>/ws/ws_xBody.asp"  	
	set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
	set rXmlObj = CreateObject("MICROSOFT.XMLDOM")
	rXmlObj.async = false
	xmlHTTP.open "POST", postURL, false
	xmlHTTp.send oXmlReg.xml  
	rv = rXmlObj.load(xmlHTTP.responseXML)		
	if not rv then
		alert "關鍵字詞傳回出現錯誤!"
		exit sub
	end if
	set xBodyNode = rXmlObj.selectSingleNode("//xBody")
	if xBodyNode.text<>"" then keywordComp xBodyNode.text
end sub

<%  if checkGIPconfig("KeywordAutoGen") then %>
sub htx_xKeyword_onChange()
	'----檢查字詞需一個字元以上
	xKeywordArray=split(reg.htx_xKeyword.value,",")
	for i=0 to ubound(xKeywordArray)
		'----去除權重符號
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr=trim(xKeywordArray(i))
		end if
		if blen(xStr)<=2 then
			alert "每一關鍵字詞長度至少二個字!"
			exit sub
		end if
	next	
	'----檢查完成
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keyword.asp?xKeyword=" & B5toUTF8(reg.htx_xKeyword.value)
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")
	if xKeywordNode.text<>"" then KeywordWTP xKeywordNode.text
end sub
<%  end if %>

sub KeywordWTP(xKeyworWTPdStr)
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xStr=""
	xReturnValue=""
	xKeywordWTPArray=split(xKeyworWTPdStr,";")
	for i=0 to ubound(xKeywordWTPArray)
            chky=msgbox("注意！"& vbcrlf & vbcrlf &"　[ "&xKeywordWTPArray(i)&" ]關鍵字詞不存在, 加入待審清單中嗎？"& vbcrlf , 48+1, "請注意！！")
            if chky=vbok then        
		xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keywordWTP.asp?xKeyword=" & B5toUTF8(xKeywordWTPArray(i))
		oXML.load(xURI)		          		
       	    end If
	next
end sub   

	clientInitForm
</script>
<script language=javascript>
function B5toUTF8(x){
	return encodeURI(x);
}
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
<%
	for each param in allModel2.selectNodes("//fieldList/field[formList!='' and inputType!='hidden']") 
	    if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") then	
		processValid param
	    end if
	next
%>
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub
<%if checkGIPconfig("MMOFolder") then %>
sub MMOFolderAdd()
	window.open "/MMO/MMOAddFolder.asp?S=P1&MMOFolderID="&reg.htx_MMOFolderID.value,"","height=400,width=550"
end sub
sub MMOFolderAddReload()
	document.location.reload()
end sub
<%end if%>
</script>

<!--#include file="DsdXMLForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1><font size=2>【目錄樹節點: <%=session("catName")%>】</font>
	<div id="Nav">
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">查詢</a>&nbsp;&nbsp;	    
	       <a href="<%=HTprogPrefix%>List.asp">回條列</a>	    
	       <%end if%>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTprogPrefix%>Add.asp?phase=add">一般資料式</a>&nbsp;&nbsp;	    
	       <a href="<%=HTprogPrefix%>Add.asp?showType=2&phase=add">URL連結式</a>&nbsp;&nbsp;	    
	       <a href="<%=HTprogPrefix%>Add.asp?showType=3&phase=add">檔案下載式</a>&nbsp;&nbsp;	    
	       <!--<a href="DsdXMLAdd_refID.asp">引用資料式</a>&nbsp;&nbsp;-->	    
	       <br>
	       <%end if%>
	<%=HTProgCap%>&nbsp;
	    <font size=2>【新增(<font color=red><%=xShowTypeStr%></font>)--主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】</td>
</div>
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
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

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

function d6date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     d6date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d6date=xy &  xm &  xd
                  end if
end function

sub doUpdateDB()
	'----940406 MMO上傳路徑/ftp等參數處理
	SQLCheck="Select rdsdcat from CtUnit C Left Join BaseDsd B ON C.ibaseDsd=B.ibaseDsd " & _
		"where C.CtUnitId='"&session("CtUnitID")&"'"
	Set RSCheck=conn.execute(SQLCheck)	
	if checkGIPconfig("MMOFolder") and RSCheck("rdsdcat")="MMO" then
		'----FTP所需參數						
		SQLM="Select MM.mmositeId+MM.mmofolderName as MMOFOlderID,MM.mmositeId,MM.mmofolderName,MS.upLoadSiteFtpip,MS.upLoadSiteFtpport,MS.upLoadSiteFtpid,MS.upLoadSiteFtppwd " & _
			" from Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
			"where MM.mmofolderID="&xUpForm("htx_MMOFolderID")
		Set RSM=conn.execute(SQLM)
		if not RSM.eof then 
			xMMOFolderID=RSM("MMOFOlderID")
	   		xFTPIPMMO = RSM("upLoadSiteFtpip")
	   		xFTPPortMMO = RSM("upLoadSiteFtpport")
	   		xFTPIDMMO = RSM("upLoadSiteFtpid")
	   		xFTPPWDMMO = RSM("upLoadSiteFtppwd")
			MMOFTPfilePath="public/"&xMMOFolderID
		end if
		'----上傳路徑
		MMOPath = session("MMOPublic") & xMMOFolderID
		if right(MMOPath,1)<>"/" then	MMOPath = MMOPath & "/"
	end if
	'----MMO上傳路徑/ftp等參數處理完成
	sql = "INSERT INTO  CuDtGeneric(showType,"
	sqlValue = ") VALUES('"+xUpForm("showType")+"'," 
	if xUpForm("showType")="4" or xUpForm("showType")="5" then
		sql = sql & "refID,"
		sqlValue = sqlValue & xUpForm("refID") & ","
	end if
	for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[formList!='']") 
		if nullText(param.selectSingleNode("fieldName")) = "xImportant" _
				 AND xUpForm("xxCheckImportant")="Y" then
			sql = sql & "xImportant,"
			sqlValue = sqlValue & pkStr(d6date(date()),",")
		elseif not ((nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") or (nullText(param.selectSingleNode("identity"))="Y")) then	
			processInsert param
		end if
	next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	'response.write sql & "<hr>"
	'response.write session("ODBCDSN")
	'response.end
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)
'	response.write xNewIdentity
'	response.end

'------- 記錄 異動 log -------- start --------------------------------------------------------	
	if checkGIPconfig("UserLogFile") then
		sql = "INSERT INTO userActionLog(loginSID,xTarget,xAction,recordNumber,objTitle) VALUES(" _
			& dfn(session("loginLogSID")) & "'0A','1'," _
			& dfn(xNewIdentity) _
			& pkstr(xUpForm("htx_sTitle"),")")
		conn.execute sql
	end if	
'------- 記錄 異動 log -------- end --------------------------------------------------------	
	
'	response.write sql & "<HR>"
	'----不同新增方式處理slave資料	
	if xUpForm("showType")="1" or xUpForm("showType")="4" or xUpForm("showType")="5" then		'----1一般資料/45引用資料
	    sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "(gicuitem,"
	    sqlValue = ") VALUES(" & dfn(xNewIdentity)
	    for each param in refModel.selectNodes("fieldList/field[formList!='']") 
		processInsert param
	    next
	    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
'response.write sql
'response.end
	    conn.Execute(SQL) 
	else	'----2URL連結/3檔案下載
	    sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "(gicuitem) Values(" & xNewIdentity & ")"
'response.write sql
'response.end
	    conn.Execute(SQL)	
	end if
	'----Email處理
	if session("CheckYN")="Y" then
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if
		set emailNode=htPageDom.selectSingleNode("SystemParameter/DsdXMLEmail")
		xEmail=nullText(emailNode)
		xEmailStr=nullText(emailNode.selectSingleNode("@xdesc"))
		fSql="Select IU.UserName,IU.Email " & _
			" from CuDtGeneric C " & _
			" Left Join CatTreeNode CTN ON C.iCtUnit=CTN.CtUnitID AND CTN.CtRootID=" & session("ItemID") & _
			" Left Join CtUserSet2 CUS2 ON CTN.CtNodeID=CUS2.CtNodeID AND CUS2.UserID=N'"&session("userID")&"' " & _
			" Left Join CtUnit CT On C.iCtUnit=CT.CtUnitID " & _
			" Left Join InfoUser IU ON CUS2.userID=IU.userID AND C.iDept=IU.deptID " & _
			" where C.iCuitem="&xNewIdentity
		'response.Write fSql
		set rs1=conn.execute(fSql)
		if not rs1.EOF then
    		    while not rs1.EOF
        	        S_Email=""""+xEmailStr+""" <"+xEmail+">"
        		R_Email=rs1(1)
            
        		Email_body="【 " & rs1(0) & " 】小姐先生 您好:" & "<br>" & "<br>" & _
                            "   現有新的待審上稿資料, 請至[後台管理網站/資料審稿作業中]審核!"& "<br><br>" & _
                            "謝謝您!"& "<br>" & "<br>" & _
                            xEmailStr & "<br>" 	
			if not isNull(rs1(1)) then                                
        			Call Send_Email(S_Email,R_Email, "上稿資料審稿通知" ,Email_body)  
		        end if 
     
		     	rs1.movenext  
		    wend
		end if	

	end if
		
	'----關鍵字詞處理
	if xUpForm("htx_xKeyword")<>"" then
	    redim iArray(1,0)
	    xStr=""
	    xReturnValue=""
	    SQLInsert=""
	    xKeywordArray=split(xUpForm("htx_xKeyword"),",")
	    weightsum=0
	    for i=0 to ubound(xKeywordArray)
	    	redim preserve iArray(1,i)
		'----分開字詞與權重符號
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
	    '----串SQL字串 
	    for k=0 to ubound(iArray,2)
	    	SQLInsert=SQLInsert+"Insert Into CuDTKeyword values("+dfn(xNewIdentity)+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if
	'----新增hyftd文章索引
	if checkGIPconfig("hyftdGIP") then
		hyftdGIPStr=hyftdGIP("add",xNewIdentity)
	end if
	'======	2006.5.8 by Gary
	if checkGIPconfig("RSSandQuery") then  
	
		SQLRSS = "SELECT YNrss FROM catTreeNode WHERE ctNodeId=" & pkStr(session("ctNodeId"),"")
		'response.write sqlrss
		'response.end
		Set RSS = conn.execute(SQLRSS)
		if not RSS.eof and RSS("YNrss")="Y" then
			session("RSS_method") = "insert"
			session("RSS_iCuItem") = xNewIdentity		
			postURL = "/ws/ws_RSSPool.asp"
			Server.Execute (postURL) 
		end if
	end if
	'======
	'----940215 Film關聯處理
	if session("FilmRelated_CorpActor")="Y" then
		if xUpForm("htx_ProductionCompanyA")<>"" then FilmRelatedStr=FilmRelated("add","Corp","製作業",xNewIdentity,xUpForm("htx_ProductionCompanyA"))
		if xUpForm("htx_DistributorsA")<>"" then FilmRelatedStr=FilmRelated("add","Corp","發行業",xNewIdentity,xUpForm("htx_DistributorsA"))
		if xUpForm("htx_Director")<>"" then FilmRelatedStr=FilmRelated("add","Actor","1",xNewIdentity,xUpForm("htx_Director"))
		if xUpForm("htx_Scriptwriter")<>"" then FilmRelatedStr=FilmRelated("add","Actor","2",xNewIdentity,xUpForm("htx_Scriptwriter"))
		if xUpForm("htx_Producer")<>"" then FilmRelatedStr=FilmRelated("add","Actor","3",xNewIdentity,xUpForm("htx_Producer"))
		if xUpForm("htx_Actor")<>"" then FilmRelatedStr=FilmRelated("add","Actor","4",xNewIdentity,xUpForm("htx_Actor"))
		if xUpForm("htx_skill")<>"" then FilmRelatedStr=FilmRelated("add","Actor","6",xNewIdentity,xUpForm("htx_skill"))
		if xUpForm("htx_Art")<>"" then FilmRelatedStr=FilmRelated("add","Actor","7",xNewIdentity,xUpForm("htx_Art"))
		if xUpForm("htx_Others")<>"" then FilmRelatedStr=FilmRelated("add","Actor","8",xNewIdentity,xUpForm("htx_Others"))
	end if
	'----940410 MMO物件引用紀錄處理
    	if checkGIPconfig("MMOFolder") then
		MMOReferendSQL=""
		if xUpForm("htx_xBody") = "" then
			MMOReferenedSQL="Delete MMOReferened where iCuItem="&pkStr(request.queryString("iCuItem"),"")
		else
			MMOReferenedSQL="Delete MMOReferened where iCuItem="&pkStr(request.queryString("iCuItem"),"")&";"
			MMOReferenedStr=xUpForm("htx_xBody")
			MMORefPos=instr(xUpForm("htx_xBody"),"MMOID=""")
			while MMORefPos > 0
			     MMOReferenedStr=mid(MMOReferenedStr,MMORefPos+7)
			     quotePos=instr(MMOReferenedStr,"""")
			     xMMOID=Left(MMOReferenedStr,quotePos-1)
			     MMOReferenedSQL=MMOReferenedSQL&"Insert Into MMOReferened values("&request.queryString("iCuItem")&","&xMMOID&");"
			     MMORefPos=instr(MMOReferenedStr,"MMOID=""")
			wend
		end if
		if MMOReferenedSQL<>"" then conn.execute(MMOReferenedSQL)
	end if
	'----950215自動繁轉簡機制
	if checkGIPconfig("convertsim") then
		session("convertsim_icuitem") = xNewIdentity
		Server.Execute ("DsdXMLAdd_convertsim.asp") 
	end if	
	
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="新增完成！"
<%if FTPErrorMSG<>"" and hyftdGIPStr="" then%>
	doneStr="新增完成！"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>"
<%elseif FTPErrorMSG="" and hyftdGIPStr<>"" then%>
	doneStr="新增完成！"+VBCRLF+VBCRLF+"<%=hyftdGIPStr%>"
<%elseif FTPErrorMSG<>"" and hyftdGIPStr<>"" then%>
	doneStr="新增完成！"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>"+VBCRLF+VBCRLF+"<%=hyftdGIPStr%>"
<%end if%>	     	
	alert(doneStr)
<%
	nextURL = HTprogPrefix & "List.asp"
	if xUpForm("nextTask") = "CuCheckList.asp?iCuItem=" then
%>
		window.open "<%=session("myWWWSiteURL")%>/ContentOnly.asp?CuItem=<%=xNewIdentity%>&ItemID=<%=session("ItemID")%>&userID=<%=session("UserID")%>"
<%		
	elseif xUpForm("nextTask") <> "" then
		nextURL = xUpForm("nextTask")  & xNewIdentity & "&op=add"
	end if
%>
	document.location.href = "<%=nextURL%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
