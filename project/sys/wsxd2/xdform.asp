<%
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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdform.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp", "ctNode", "cuItem", "xItem", "spec", "CtUnit")
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
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->	
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function message(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	message=""
  	exit function
  elseif instr(1,xs,"<P",0)>0 or instr(1,xs,"<BR",0)>0 or instr(1,xs,"<td",0)>0 then
 	message=xs
  	exit function
  end if
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR/>")
  	message = replace(xs,chr(10),"<BR/>")
end function

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	

	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
	xv = htPageDom.load(LoadXML)
		if htPageDom.parseError.reason <> "" then 
  		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
  		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
  		Response.End()
		end if

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	
	if request.queryString("ctNode") = "" and request.queryString("cuItem") <> "" then
		sql = "SELECT  t.ctNodeID" _
			& " FROM CuDTGeneric AS htx JOIN CatTreeNode AS t " _
			& " ON t.ctUnitID=htx.iCtUnit" _
			& " WHERE htx.iCuItem=" & pkstr(request("cuItem"),"")
		set RS = conn.execute(sql)
		ctNode = RS(0)
		xItem = request("cuItem")
	else
		ctNode = request("ctNode")
		xItem = request("xItem")
	end if
	
	
	if (not isNumeric(ctNode)) then 
		response.end
	end if
	if (not isNumeric(xItem)) then 
		response.end
	end if

	

	myTreeNode = ctNode
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>" & VBCRLF
	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>" & VBCRLF
	response.write "<mp>"&request("mp")&"</mp>" & VBCRLF

	if xItem <> "" then
		sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(ctNode,"")
		set RS = conn.execute(sql)
		xRootID = RS("CtRootID")
		xCtUnitName = RS("CatName")
		xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
		xParent = RS("DataParent")
		myxslList = RS("xslData")
		response.write "<xslData>"&myxslList&"</xslData>" & VBCRLF
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
		doCP
	end if
	response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>" & VBCRLF
	
	for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
		processXDataSet
	next	
	
	if isNull(xRootID) then _
		xRootID=nullText(refModel.selectSingleNode("MenuTree"))

	set htPageDom1 = Server.CreateObject("MSXML2.DOMDocument.3.0")
	htPageDom1.setProperty "SelectionLanguage", "XPath"	
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\" & request("spec") & ".xml"
	
	Fxv = htPageDom1.load(LoadXML)

	if htPageDom.parseError.reason <> "" then 
		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
		Response.End()
	end if

	set FrefModel = htPageDom1.selectSingleNode("//DataSchemaDef/language[mp=" & request("mp") & "]")
	'response.write "<test>" & FrefModel.xml & "</test>"
	set list = FrefModel.selectSingleNode("FieldList")
	'set list = FrefModel.selectSingleNode("//fieldList")
	
'	if request.queryString("ctNode") = "" then
'		sql = "SELECT  t.ctNodeID" _
'			& " FROM CuDTGeneric AS htx JOIN CatTreeNode AS t " _
'			& " ON t.ctUnitID=htx.iCtUnit" _
'			& " WHERE htx.iCuItem=" & pkstr(request("cuItem"),"")
'		set RS = conn.execute(sql)
'		ctNode = RS(0)
'		xItem = request("cuItem")
'	else
'		ctNode = request("ctNode")
'		xItem = request("xItem")
'	end if

	myTreeNode = ctNode
	if request("CtUnit") <> "" then
		response.write "<CtUnit>"&request("CtUnit")&"</CtUnit>" & VBCRLF
	end if
	response.write "<xItem>"&request("xItem")&"</xItem>" & VBCRLF
	response.write "<ctNode>"&request("ctNode")&"</ctNode>" & VBCRLF	
	response.write "<PageTitle>" & nullText(FrefModel.selectSingleNode("PageTitle")) & "</PageTitle>" & VBCRLF
	response.write "<TopInfo>" & nullText(FrefModel.selectSingleNode("TopInfo")) & "</TopInfo>" & VBCRLF
	response.write "<Action>" & nullText(FrefModel.selectSingleNode("Action")) & "</Action>" & VBCRLF	
	set script = FrefModel.selectSingleNode("SCRIPT")
	response.write script.xml
	response.write list.xml
	
	'response.write "<FieldList>" & VBCRLF
	'for each field in list.selectNodes("//Field")
	'	response.write "<Field>" & VBCRLF
	'	response.write "	<FieldTitle>" & nullText(field.selectSingleNode("Title")) & "</FieldTitle>" & VBCRLF
	'	response.write "	<Type>" & nullText(field.selectSingleNode("Type")) & "</Type>" & VBCRLF
	'	response.write "	<Name>" & nullText(field.selectSingleNode("Name")) & "</Name>" & VBCRLF		
	'	response.write "	<Value>" & nullText(field.selectSingleNode("Value")) & "</Value>" & VBCRLF
	'	response.write "	<Size>" & nullText(field.selectSingleNode("Size")) & "</Size>" & VBCRLF
	'	response.write "	<Cols>" & nullText(field.selectSingleNode("Cols")) & "</Cols>" & VBCRLF
	'	response.write "	<Rows>" & nullText(field.selectSingleNode("Rows")) & "</Rows>" & VBCRLF
	'	response.write "</Field>" & VBCRLF
	'next
	'response.write "</FieldList>" & VBCRLF
	
	set htPageDom1 = Server.CreateObject("MSXML2.DOMDocument.3.0")
	htPageDom1.setProperty "SelectionLanguage", "XPath"
	htPageDom1.async = false
	htPageDom1.setProperty("ServerHTTPRequest") = true		
	LoadXML1 = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\button.xml"	
	xv1 = htPageDom1.load(LoadXML1)

	set B_submit = htPageDom1.selectSingleNode("//submit/language[mp=" & request("mp") & "]")
	L_submit = nullText(B_submit.selectSingleNode("label"))
	set B_reset = htPageDom1.selectSingleNode("//reset/language[mp=" & request("mp") & "]")
 	L_reset = nullText(B_reset.selectSingleNode("label"))
  
%>
<pHTML>
<input type="submit" value ="<%=L_submit%>" class="Button" /> 
<input type="reset" value ="<%=L_reset%>" class="Button" />	
</pHTML>
<%	
	
sub doCP
	sql = "SELECT  htx.*, xr1.deptName, u.CtUnitName " _
		& ", (SELECT count(*) FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y' AND dhtx.xiCUItem=htx.iCuItem) AS attachCount " _
		& ", (SELECT count(*) FROM CuDTPage AS phtx" _
		& " WHERE bList='Y' AND phtx.xiCUItem=htx.iCuItem) AS pageCount " _
		& " FROM CuDtGeneric AS htx LEFT JOIN dept AS xr1 ON xr1.deptID=htx.idept" _
		& " LEFT JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " WHERE htx.iCuItem=" & pkstr(xItem,"")

	set RS = conn.execute(sql)
	if not RS.eof then

	cPostDate = d7date(RS("xPostDate"))
	scPostDate = "中華民國 " & mid(cPostDate,1,2) & " 年 " & mid(cPostDate,4,2) & " 月 " & mid(cPostDate,7,2) & " 日" 
%>
		<MainArticle iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">
			<Caption><%=RS("sTitle")%></Caption>
			<Content><![CDATA[<%=message(RS("xBody"))%>]]></Content>
		</MainArticle>
<%
	end if
end sub	
%>
<!--#include file="x1Menus.inc" -->
<%'======	2006.4.27 	
			set htPageDom1 = Server.CreateObject("MSXML2.DOMDocument.3.0")
			htPageDom1.setProperty "SelectionLanguage", "XPath"
			
			htPageDom1.async = false
			htPageDom1.setProperty("ServerHTTPRequest") = true		
			LoadXML1 = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\button.xml"	
			xv1 = htPageDom1.load(LoadXML1)
			
			set B_path = htPageDom1.selectSingleNode("//path/language[mp=" & request("mp") & "]")
			L_path = nullText(B_path.selectSingleNode("label"))							
			set B_home = htPageDom1.selectSingleNode("//home/language[mp=" & request("mp") & "]")
			L_home = nullText(B_home.selectSingleNode("label"))	
%>
	<xxpath><%=L_path%></xxpath>	
	<xxhome><%=L_home%></xxhome>
</hpMain>