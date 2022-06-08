<%@ CodePage = 65001 %><%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:37
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdQP1.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp", "ctNode", "a", "CtNode", "CtUnit", "BaseDSD")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/site/cdic/inc/dbFunc.inc" -->
<!--#Include virtual = "/site/cdic/inc/xDataSet.inc" -->
<!--#Include virtual = "/site/cdic/inc/htUIGen.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

Dim RSKey

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/cdic/GipDSD") & "\xdmp" & request("mp") & ".xml"
'		response.write LoadXML & "<HR>"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	myTreeNode = request("ctNode")
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"

	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(request("ctNode"),"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
	myxslList = RS("xslData")
	response.write "<xslData>"&myxslList&"</xslData>"
	myParent = xParent
	xLevel = RS("DataLevel") - 1
	if RS("CtNodeKind") <> "C" then
		xLevel = xLevel -1
		myTreeNode = xParent
	end if
	upParent = 0
	
	while xParent<>0
		sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		if RS("DataLevel") = xLevel then	upParent = xParent
		xPathStr = "<xPathNode Title=""" & deAmp(RS("CatName")) & """ xNode=""" & RS("ctNodeID") & """ />" & xPathStr
		xParent = RS("DataParent")
	wend
	response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>"
	if request.querystring("a")="np" then
	    qURL = "np.asp?CtNode="&request.queryString("CtNode")
	else
	     qURL = "lp.asp?CtNode="&request.queryString("CtNode")&"&amp;CtUnit="&request.queryString("CtUnit") _
		& "&amp;BaseDSD=" & request.queryString("BaseDSD")
	end if
%>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" />
	<xxxslData>xQueryForm</xxxslData>
<%  	
		doFP
	
  for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
	processXDataSet
  next



function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

sub doFP
'-------準備前端呈現需要呈現欄位的DSD xmlDOM
	SQLTable="Select sBaseTableName from BaseDSD where iBaseDSD=" & pkStr(request.queryString("BaseDSD"),"")
	Set RSTable=conn.execute(SQLTable)		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/cdic/GipDSD/CtUnitX" & request.querystring("CtUnit") & ".xml")
    	if fso.FileExists(filePath) then
    		LoadXMLDSD = server.MapPath("/site/cdic/GipDSD/CtUnitX" & request.querystring("CtUnit") & ".xml")
    	else
    		LoadXMLDSD = server.MapPath("/site/cdic/GipDSD/CuDTx" & request.querystring("BaseDSD") & ".xml")
    	end if 
	set DSDDom = Server.CreateObject("MICROSOFT.XMLDOM")
	xv = DSDDom.load(LoadXMLDSD)
'	response.write xv & "<HR>"
  	if DSDDom.parseError.reason <> "" then 
    		Response.Write("DSDDom parseError on line " &  DSDDom.parseError.line)
    		Response.Write("<BR>Reason: " &  DSDDom.parseError.reason)
    		Response.End()
  	end if	
    	set root = DSDDom.selectSingleNode("DataSchemaDef") 	
    	'----Load XSL樣板
    	set oxsl = server.createObject("microsoft.XMLDOM")
   	oxsl.async = false
   	xv = oxsl.load(server.mappath("/site/cdic/GipDSD/CtUnitXOrder.xsl"))   
    	'----複製Slave的dsTable,並依順序轉換
	set DSDNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&RSTable(0)&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement  
	for each param in nxmlnewNode.selectNodes("field[queryListClient='']") 
		set romoveNode=nxmlnewNode.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode.removeChild romoveNode
	next    	  
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")    	
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&RSTable(0)&"']")
    	'----複製CuDTGeneric的dsTable,並依順序轉換
    	set GenericNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDTGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement
	for each param in nxmlnewNode2.selectNodes("field[queryListClient='']") 
		set romoveNode=nxmlnewNode2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode2.removeChild romoveNode
	next        	    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDTGeneric']")       	
  	set DSDrefModel = DSDDom.selectSingleNode("//dsTable")
  	set DSDallModel = DSDDom.selectSingleNode("//DataSchemaDef")
	response.write DSDallModel.xml
%>
		<pHTML>
<form method="POST" name="reg" action="<%=qURL%>">
<INPUT TYPE="hidden" name="submitTask" value="" />
<object data="/inc/calendar.htm" id="calendar1" type="text/x-scriptlet" width="245" height="160" style="position:absolute;top:0;left:230;visibility:hidden"></object>
<INPUT TYPE="hidden" name="CalendarTarget" />
<CENTER>
<table width="100%"  border="0" cellspacing="0" cellpadding="0" summary="資料表格查詢">
  <caption align="left">
  <%=xCtUnitName%>查詢
  </caption>
<%	for each param in DSDallModel.selectNodes("//fieldList/field[queryListClient!='']")
		response.write "<tr><th scope='row'>"
		response.write nullText(param.selectSingleNode("fieldLabel")) & "</th>" & vbCRLF
		response.write "<td>"
			processQParamField param
		response.write "</td></tr>" & vbCRLF
	next
%>	
</table>
<br/>
            <input type="submit" value ="查　詢" class="cbutton" />
            &amp;nbsp;　&amp;nbsp;
            <input type="reset" value ="重　填" class="cbutton" />
</CENTER>
</form>
<script language="vbs">
<![CDATA[
Dim CanTarget
Dim followCanTarget

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
]]>
</script>
<noscript>
很抱歉您的瀏覽器不支援JavaScript，頁面切換Script不影響到畫面瀏覽
</noscript>
		</pHTML>
<%		
end sub
%>
<!--#include file="x0Menus.inc" -->
<%
	dim weekdayStr(6)
	weekdayStr(0)="日":weekdayStr(1)="一":weekdayStr(2)="二":weekdayStr(3)="三":weekdayStr(4)="四":weekdayStr(5)="五":weekdayStr(6)="六"
	dateStr_c="今天是民國"&year(date())-1911&"年"&month(date())&"月"&day(date())&"日　星期"&weekdayStr(weekday(date())-1)
	response.write "<Today_C>"&dateStr_c&"</Today_C>"
	response.write "<upDate>"&d7date(session("upDate"))&"</upDate>"
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	Response.Write "<Counter>" & rs("counts") & "</Counter>"
%>
</hpMain>
