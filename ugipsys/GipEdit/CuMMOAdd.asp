<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="多媒體物件選取"
HTProgFunc="新 增"
HTProgCode="GC1AP1"
HTProgPrefix="CuMMO" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function


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

cq=chr(34)
ct=chr(9)
cl="<" & "%"
cr="%" & ">"
formFunction = "add"    

'----檔案目錄處理
if request.querystring("phase")<>"add" and request.querystring("submitTask")<>"SHOW" then
    apath = server.mappath(session("MMOPublic"))
    Set xup = Server.CreateObject("TABS.Upload")
    xup.codepage=65001
    xup.Start apath
else
    Set xup = Server.CreateObject("TABS.Upload")
end if

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function	    

if xUpForm("submitTask")="ADD" then
	'----940406 MMO上傳路徑/ftp等參數處理
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
	'----MMO上傳路徑/ftp等參數處理完成
        '----開始處理form相關
        set htPageDom = session("codeXMLSpec_MMO")
        set allModel2 = session("codeXMLSpec2_MMO")  
        set refModel = htPageDom.selectSingleNode("//dsTable")
        set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
    
  	'----資料庫處理
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	'----新增主表
	sql = "INSERT INTO  CuDtGeneric("
	sqlValue = ") VALUES("
	for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[formList!='']") 
		if nullText(param.selectSingleNode("fieldName")) = "ximportant" _
				 AND xUpForm("xxCheckImportant")="Y" then
			sql = sql & "xImportant,"
			sqlValue = sqlValue & pkStr(d6date(date()),",")
		elseif nullText(param.selectSingleNode("identity"))<>"Y" then	
			processInsert param
		end if
	next
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)
'------- 記錄 異動 log -------- start --------------------------------------------------------	
	if checkGIPconfig("UserLogFile") then
		sql = "INSERT INTO userActionLog(loginSID,xTarget,xAction,recordNumber,objTitle) VALUES(" _
			& dfn(session("loginLogSID")) & "'0A','1'," _
			& dfn(xNewIdentity) _
			& pkstr(xUpForm("htx_sTitle"),")")
		conn.execute sql
	end if	
'------- 記錄 異動 log -------- end --------------------------------------------------------	

	'----新增明細表
	    sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "(giCuItem,"
	    sqlValue = ") VALUES(" & dfn(xNewIdentity)
	    for each param in refModel.selectNodes("fieldList/field[formList!='']") 
		processInsert param
	    next
	    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	    conn.Execute(SQL) 
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
	    	SQLInsert=SQLInsert+"Insert Into CuDtkeyword values("+dfn(xNewIdentity)+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if	
	response.write "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/></head>"	
	response.write "<script language='javascript'>"
	response.write "alert('新增完成！');"
	response.write "window.navigate('CuMMOQuery.asp');"
	response.write "</script>"
	response.write  "</html>"	    
%>
	<%'response.end	
elseif request("submitTask")="SHOW" then
    SQL="Select ctUnitId,ibaseDsd " & _
    	"from CtUnit where ctUnitId=" & pkstr(request("htx_CtUnitID"),"")
    Set RSM=conn.execute(SQL)
    if not isNull(RSM(0)) then
    	xCtUnitID=RSM(0)
    	xiBaseDSD=RSM(1)
    	session("CtUnitID") = xCtUnitID
    	session("iBaseDSD") = xiBaseDSD   
    	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(xCtUnitID) & ".xml")    	
    end if
    Set fso = server.CreateObject("Scripting.FileSystemObject")
    if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
    end if 
    
    '----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
    Set fso = server.CreateObject("Scripting.FileSystemObject")
    set oxml = server.createObject("microsoft.XMLDOM")
    oxml.async = false
    oxml.setProperty "ServerHTTPRequest", true
    
    xv = oxml.load(LoadXML)
    if oxml.parseError.reason <> "" then
	Response.Write("XML parseError on line " &  oxml.parseError.line)
	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
	Response.End()
    end if
    set session("codeXMLSpec_MMO") = oxml
    '----Load XSL樣板
    set oxsl = server.createObject("microsoft.XMLDOM")
    oxsl.async = false
    xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))          
    '----混合field順序
    set nxml0 = server.createObject("microsoft.XMLDOM")
    nxml0.LoadXML(oxml.transformNode(oxsl))
    set session("codeXMLSpec2_MMO") = nxml0  
    '----開始處理form相關
    set htPageDom = session("codeXMLSpec_MMO")
    set allModel2 = session("codeXMLSpec2_MMO")  
    set refModel = htPageDom.selectSingleNode("//dsTable")
    set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
end if
%>

<title>新增多媒體物件</title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<link href="css/editor.css" rel="stylesheet" type="text/css">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="PopFormName">多媒體物件新增</div>
<%if request("submitTask")="SHOW" then%>
<!--#include file="CuMMOForm.inc"-->
<%else%>
<form action="" method="" id="PopForm">
<INPUT TYPE=hidden name=submitTask value="">
<table cellspacing="0">
  <tr>
    <td class="Label">主題單元</td>
    <td>
	<Select name="htx_CtUnitID" size=1 onchange="VBS:MMOFolderIDSelect()">
	<option value="">請選擇</option>
		<%SQL="Select ctUnitId,ctUnitName from CtUnit CT Left Join baseDsd B ON Ct.ibaseDsd=B.ibaseDsd where rdsdcat='MMO'"
		SET RSS=conn.execute(SQL)
		While not RSS.EOF%>

		<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
		<%	RSS.movenext
		wend%>
	</select>
      </td>
    <td class="Label">物件存放目錄</td>
    <td>
	<Select name="htx_MMOFolderID_1" size=1>
	</select>
    </td>
  </tr>
</table>
<hr>
	<input type="button" class="InputButton" value="確定" onClick="formShowSubmit()">
	<input type="button" class="InputButton" value="回查詢" onClick="Javascript:window.navigate('CuMMOQuery.asp');">
</form>
<%end if%>
</body>
</html>
<script language=vbs>
sub formShowSubmit()
	if PopForm.htx_MMOFolderID_1.value="" then
		alert "請選擇主題單元後, 再選擇物件存放目錄!"
		PopForm.htx_MMOFolderID_1.focus
		exit sub
	end if
        PopForm.submitTask.value="SHOW"
        PopForm.Submit
end Sub

sub MMOFolderIDSelect()
 	set xsrc = document.all("htx_MMOFolderID_1")
 	removeOption(xsrc)
 	set oXML = createObject("Microsoft.XMLDOM")
 	oXML.async = false
 	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_MMOFolderID.asp?CtUnitID=" & PopForm.htx_CtUnitID.value
 	oXML.load(xURI)
 	set pckItemList = oXML.selectNodes("divList/row")
 	for each pckItem in pckItemList
  		xaddOption xsrc, pckItem.selectSingleNode("mValue").text,pckItem.selectSingleNode("mCode").text
 	next
 	xsrc.selectedIndex = 0
end sub	
sub xaddOption(xlist,name,value)
 	set xOption = document.createElement("OPTION")
 	xOption.text=name
 	xOption.value=value
 	xlist.add(xOption)
end sub
sub removeOption(xlist)
 	for i=xlist.options.length-1 to 0 step -1
  		xlist.options.remove(i)
 	next
 	xlist.selectedIndex = -1
end sub

'----------------------------------
sub resetForm
	PopForm.reset()
	reg.reset
	clientInitForm	
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
<%if request("submitTask")="SHOW" then%>
	reg.htx_MMOFolderID.value="<%=request("htx_MMOFolderID_1")%>"
<%	for each xparam in allModel.selectNodes("//fieldList/field[formList!='']") 
	    if not (nullText(xparam.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") then
		AddProcessInit xparam
	    end if
	next
	if nullText(allModel.selectSingleNode("//fieldList/field[fieldName='topCat']/showList"))<>"" then _
		response.write "reg.htx_topCat.value = """ & left(session("tDataCat"),1) & """" 
  end if
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
<%if request("submitTask")="SHOW" then%>
Sub formAddSubmit()	
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
	mMsg = "「{0}」欄位格式不符合！"	
<%
	for each param in allModel.selectNodes("//fieldList/field[formList!='' and inputType!='hidden']") 
		processValid param
	next
%>
        reg.submitTask.value="ADD"
        reg.Submit
End Sub
sub MMOFolderAdd()
	window.open "/MMO/MMOAddFolder.asp?S=P1&MMOFolderID="&reg.htx_MMOFolderID.value,"","height=400,width=550"
end sub
sub MMOFolderAddReload()
	document.location.reload()
end sub
<%end if%>
</script>

