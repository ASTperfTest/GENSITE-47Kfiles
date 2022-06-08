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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdCP2.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("ctNode", "ctunit", "mp", "cuItem", "xItem", "memID", "gstyle", "xq_xCat", "htx_vgroup", "pagesize", "nowPage")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
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
%><% response.contentType="text/xml" %><?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/MSclient.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<!--#Include file = "time.inc" -->
<!--#Include virtual = "/inc/GSSUtility.inc" -->
<%'即時字詞顯示include %>
<!--#Include virtual = "/inc/ReplaceAndFindKeyword.inc" -->
<%
	if (not isNumeric(request("ctNode"))) then 
		response.end
	end if
	if (not isNumeric(request("ctunit"))) then 
		response.end
	end if
	mp = request("mp")
	if ( Instr(mp, "<") > 0 or Instr(mp, ">") > 0 or Instr(mp, "'") > 0 ) then
		response.end
	end if
%>
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
	
	'LoadXML = server.MapPath("GipDSD") & "\xdmp" & request("mp") & ".xml"
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
	'response.write LoadXML & "<HR>"
	xv = htPageDom.load(LoadXML)
	'response.write xv & "<HR>"
  if htPageDom.parseError.reason <> "" then 
  	Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
  	Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
  	Response.End()
  end if

  set refModel = htPageDom.selectSingleNode("//MpDataSet")

	if request.queryString("ctNode") = "" then
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

	'把count移到viewcounter.aspx 
	'sql = "UPDATE CuDTGeneric SET ClickCount = ClickCount + 1 WHERE iCUItem = '" & xItem & "' "
	'conn.execute(sql)

	myTreeNode = ctNode
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"
	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"
	
	'-----會員登入區塊-----GSSUtility
	call GetLoginInfo(request("memID"), request("gstyle")) 

	'---start---加入推薦詞彙---檢查是否開啟---2008/09/10---vincent---	
	Response.Write "<commendword>" 
	sql = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'commendWord') AND (mCode = 'inter')"
	Set crs = conn.execute(sql)
	if not crs.eof then
		if crs("mValue") = "1" then
			response.write "<isopen>Y</isopen>"
		else
			response.write "<isopen>N</isopen>"
		end if
	else
		response.write "<isopen>N</isopen>"
	end if
	crs.close
	set crs = nothing
	Response.Write "</commendword>" 
	'---end---加入推薦詞彙---檢查是否開啟---2008/09/10---vincent---
%>
<!--#include file="attributeVote.inc" -->	
<!--#include file="relatedDocument.inc" -->	
<%	
	'---抓出此節點所使用的版型---
	sql = "SELECT * FROM CatTreeNode where CtNodeID = " & pkstr(ctNode,"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
	myxslList = RS("xslData")
	response.write "<xslData>" & myxslList & "</xslData>"
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
	
	'---特殊CP版型---vincent---2008/12/11---
	sql = "SELECT KMcat FROM CuDTGeneric WHERE icuitem = " & pkstr(xItem,"")
	set cprs = conn.execute(sql)
	if not cprs.eof then
		response.write "<cpStyle>" & cprs("KMcat") & "</cpStyle>"
	else
		response.write "<cpStyle></cpStyle>"
	end if
	cprs.close
	set cprs = nothing
	'20090617 新優質農家 產銷專欄 農業生活 導覽路徑秀資料分類
	'http://kmw.hyweb.com.tw/ct.asp?xItem=82412&ctNode=1582&mp=1
	'http://kmw.hyweb.com.tw/lp.asp?CtNode=1582&CtUnit=2766&BaseDSD=7&mp=1&xq_xCat=02
	sql1 = "SELECT * FROM CuDTGeneric WHERE iCUItem = '"&xItem&"'"
	set rs1 = conn.execute(sql1)
	if not rs1.eof and rs1("NewtopCat") <> "" then '新優質農家 產銷專欄
		if request("ctNode") = "1582" then
			sql2="SELECT * FROM CodeMain WHERE (codeMetaID = N'new_article') AND (mCode = '"&rs1("NewtopCat")&"')"
			set rs2 = conn.execute(sql2)
		end if
		if request("ctNode") = "1583" then
			sql2="SELECT * FROM CodeMain WHERE (codeMetaID = N'new_topcat') AND (mCode = '"&rs1("NewtopCat")&"')"
			set rs2 = conn.execute(sql2)
		end if
		response.write "<CatList>"
		response.write "<mediapath><mcode>"&rs1("NewtopCat")&"</mcode>"
		response.write "<mvalue>"&rs2("mValue")&"</mvalue>"
		response.write "<murl>lp.asp?CtNode="&request("CtNode")&"&amp;CtUnit="&rs1("iCTUnit")&"&amp;BaseDSD="&rs1("iBaseDSD")&"&amp;xq_xCat="&rs1("NewtopCat")&"</murl>"
		response.write "</mediapath>"
		response.write "</CatList>"
	end if
	if not rs1.eof and request("ctNode") = "1581" then '農業生活
		sql2="SELECT * FROM CodeMain WHERE (codeMetaID = N'new_life') AND (mCode = '"&rs1("topCat")&"')"
		set rs2 = conn.execute(sql2)
		response.write "<CatList>"
		response.write "<mediapath><mcode>"&rs1("topCat")&"</mcode>"
		response.write "<mvalue>"&rs2("mValue")&"</mvalue>"
		response.write "<murl>lp.asp?CtNode="&request("CtNode")&"&amp;CtUnit="&rs1("iCTUnit")&"&amp;BaseDSD="&rs1("iBaseDSD")&"&amp;xq_xCat="&rs1("topCat")&"</murl>"
		response.write "</mediapath>"
		response.write "</CatList>"
	end if
	'20090617 新優質農家 產銷專欄 農業生活 導覽路徑秀資料分類end
	'// 2008/08/19 update.
	'// 丟出 aliu 要的被關聯的 第二組資料大類 ........(start)
	xTopCat = request("xq_xCat")
	xVgroup = request("htx_vgroup")
	response.write "<CatList>"
	if request("ctNode") = "1584" then
		dim caturl : caturl = ""
		sql = "SELECT CatTreeNode.CtNodeID, CtUnit.CtUnitID, CtUnit.iBaseDSD FROM CatTreeNode INNER JOIN " & _
			"CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID WHERE (CatTreeNode.CtNodeID = 1584)"
		set urlrs = conn.execute(sql)
		if not urlrs.eof then
			if xVgroup <> "" then
				caturl = "/lp.asp?ctNode=" & urlrs("CtNodeID") & "&amp;CtUnit=" & urlrs("CtUnitID") & "&amp;BaseDSD=" & urlrs("iBaseDSD") & "&amp;xq_xCat=" & xTopCat
			end if
		end if
		urlrs.close
		if xTopCat <> "" Then
			sql = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'mediacata') AND (mCode = '" & xTopCat & "')"
			set catrs = conn.execute(sql)
			if not catrs.eof then
				response.write "<mediapath><mcode>" & catrs("mCode") & "</mcode><mvalue>" & catrs("mValue") & "</mvalue><murl><![CDATA[" & caturl & "]]></murl></mediapath>"
			else
				response.write "<mediapath><mcode></mcode><mvalue></mvalue><murl></murl></mediapath>"
			end if
			catrs.close
			'catrs = nothing
			if xVgroup <> "" then
				sql = "Select iCuItem, sTitle from CuDTGeneric where topCat = '" & xTopCat & "' and ictunit=1988 and icuitem = '" & xVgroup & "'"
				set cat2rs = conn.execute(sql)
				if not cat2rs.eof then
					response.write "<mediapath><mcode>" & cat2rs("icuitem") & "</mcode><mvalue>" & cat2rs("stitle") & "</mvalue><murl></murl></mediapath>"
				else
					response.write "<mediapath><mcode></mcode><mvalue></mvalue><murl></murl></mediapath>"
				end if
				cat2rs.close
			else
				response.write "<mediapath><mcode></mcode><mvalue></mvalue><murl></murl></mediapath>"
			end if
		else
			'---沒有參數.自動幫它帶出---
			sql = " SELECT CodeMain.mCode, CodeMain.mValue, CuDTGeneric_1.iCUItem, CuDTGeneric_1.sTitle, CuDTGeneric.iBaseDSD, CuDTGeneric.iCTUnit " & _
						" FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN CuDTGeneric AS CuDTGeneric_1 " & _
						" ON CuDTGeneric.vGroup = CuDTGeneric_1.iCUItem WHERE (CuDTGeneric.iCUItem = " & request("xItem") & ") " & _
						" AND (CuDTGeneric_1.iCTUnit = 1988) AND (CodeMain.codeMetaID = 'mediacata') "
			
			set catrs = conn.execute(sql)
			if not catrs.eof then
				caturl = "/lp.asp?ctNode=1584&amp;CtUnit=" & catrs("iCTUnit") & "&amp;BaseDSD=" & catrs("iBaseDSD") & "&amp;xq_xCat=" & catrs("mCode") & "&amp;htx_vgroup=" & catrs("iCUItem")
				response.write "<mediapath><mcode>" & catrs("mCode") & "</mcode><mvalue>" & catrs("mValue") & "</mvalue><murl>" & caturl & "</murl></mediapath>"
				response.write "<mediapath><mcode>" & catrs("iCUItem") & "</mcode><mvalue>" & catrs("sTitle") & "</mvalue><murl></murl></mediapath>"
			end if
			catrs.close
			set catrs = nothing
		end if
	else
		response.write "<mediapath><mcode></mcode><mvalue></mvalue><murl></murl></mediapath>"
		response.write "<mediapath><mcode></mcode><mvalue></mvalue><murl></murl></mediapath>"
	end if
	response.write "</CatList>"
	'// 丟出 aliu 要的被關聯的 第二組資料大類 ........(end)
	'-----知識拚圖的AP Block-----
%>
	<!--#include file="KnowledgeJigsaw.inc" -->	
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" />
<%  	
'-------準備前端呈現需要呈現欄位的DSD xmlDOM
	SQLTable="Select BD.sBaseTableName,CG.iBaseDSD,CG.iCTUnit " & _
		"from CuDtGEneric CG Left Join BaseDSD BD ON CG.iBaseDSD=BD.iBaseDSD " & _
		"where CG.iCUItem=" & pkStr(xItem,"")
	Set RSTable=conn.execute(SQLTable)		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & RSTable("iCTUnit") & ".xml")
    	if fso.FileExists(filePath) then
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & RSTable("iCTUnit") & ".xml")
    	else
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & RSTable("iBaseDSD") & ".xml")
    	end if  
	set DSDDom = Server.CreateObject("MICROSOFT.XMLDOM")
	DSDDom.async = false
	DSDDom.setProperty("ServerHTTPRequest") = true	
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
   	xv = oxsl.load(server.mappath("/site/" & session("mySiteID") & "/GipDSD/CtUnitXOrder.xsl"))   
    	'----複製Slave的dsTable,並依順序轉換
	set DSDNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&RSTable(0)&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement  
	for each param in nxmlnewNode.selectNodes("field[formListClient='']") 
		set romoveNode=nxmlnewNode.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode.removeChild romoveNode
	next    	  
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")    	
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&RSTable(0)&"']")
    	'----複製CuDtGeneric的dsTable,並依順序轉換
    	set GenericNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement
	for each param in nxmlnewNode2.selectNodes("field[(formListClient='' and fieldName!='stitle') or inputType='hidden']") 
		set romoveNode=nxmlnewNode2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode2.removeChild romoveNode
	next        	    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDtGeneric']")       	
  	set DSDrefModel = DSDDom.selectSingleNode("//dsTable")
  	set DSDallModel = DSDDom.selectSingleNode("//DataSchemaDef")
  	'----混合field順序
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(DSDDom.transformNode(oxsl))

    'Modified By Leo  2011-07-13       回列表頁、往上一篇、往下一篇    Start
	if myxslList<>"" then
		doFP
        setPageURL_CuDTx7
	else
		doCP
        setPageURL_CuDTxT
	end if
	'Modified By Leo  2011-07-13       回列表頁、往上一篇、往下一篇     End

	for each xBlock in DSDallModel.selectNodes("refDataBlock")
		dorefDataGroup xBlock
	next

'	for each xBlock in DSDallModel.selectNodes("refDataBlock")
'		dorefDataBlock xBlock
'	next

  if RS("attachCount") > 0 then
	fSql = "SELECT dhtx.*" _
		& " FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y'" _
		& " AND dhtx.xiCUItem=" & pkStr(RS("iCUItem"),"") _
		& " ORDER BY dhtx.listSeq, dhtx.aEditDate DESC"
	set RSlist = conn.execute(fSql)
	Dim picCount : picCount = 0
	'ivy
	sessionUrl = "http://kmwebsys.coa.gov.tw/site/coa"
	response.write "<AttachmentList>" & vbCRLF
	while not RSlist.eof
		'response.write "<Attachment><URL><![CDATA[public/Attachment/" & RSlist("nfileName") _
		'& "]]></URL><Caption><![CDATA[" & RSlist("atitle") & "]]></Caption><Attachkind><![CDATA[" & RSlist("AttachkindA") & "]]></Attachkind><Attachtype><![CDATA[" & RSlist("Attachtype") & "]]></Attachtype></Attachment>"
		response.write "<Attachment>"
		
		response.write "<IsImageFile>"
		select case LCASE(right( RSlist("nfileName"), 4))
			case ".gif",".png",".jpg",".bmp","jpge"
				response.write "Y"
			case else
				response.write "N"
		end select
		response.write "</IsImageFile>"
		
		response.write "<URL><![CDATA[public/Attachment/" & RSlist("nfileName") & "]]></URL>"
		response.write "<ShowPictureURL><![CDATA[" & sessionUrl & "/wsxd2/ShowPicture.aspx?fileName=" & RSlist("nfileName") & "&CuItem=" & request("xItem") & "&PathType=A" & "]]></ShowPictureURL>"
		response.write "<Caption><![CDATA[" & RSlist("atitle") & "]]></Caption>"
		response.write "<Descxx><![CDATA[" & RSlist("aDesc") & "]]></Descxx>"
		fileType = LCase(mid(RSlist("nfileName"), Instr(RSlist("nfileName"), ".") + 1, len(RSlist("nfileName")) - Instr(RSlist("nfileName"), ".") ) )
		response.write "<fileType>" & fileType & "</fileType>"
		if Instr(allFileType, fileType) = 0 then allFileType = allFileType & fileType & ";" 
		if fileType = "jpg" or fileType = "gif" then
			picCount = picCount + 1
			response.write "<picCount>" & picCount & "</picCount>"
		else
			response.write "<picCount></picCount>"
		end if		
		response.write "</Attachment>"
		RSlist.moveNext
	wend
	response.write "<allFileType><![CDATA[" & allFileType & "]]></allFileType>"
	response.write "</AttachmentList>"
  end if

  if RS("pageCount") > 0 then
	fSql = "SELECT dhtx.*, n.*" _
		& " FROM CuDTPage AS dhtx" _
		& " JOIN CuDtGeneric AS n ON dhtx.nPageID=n.iCuItem" _
		& " WHERE bList='Y'" _
		& " AND dhtx.xiCUItem=" & pkStr(RS("iCUItem"),"") _
		& " ORDER BY dhtx.listSeq"
	set RSlist = conn.execute(fSql)
	response.write "<ReferenceList>" & vbCRLF
	while not RSlist.eof
		response.write "<Reference iCuItem='" & RSlist("nPageID") & "'><URL>content.asp?CuItem=" & RSlist("nPageID")& "&amp;mp=" & request("mp") _
			& "</URL><Caption><![CDATA[" & RSlist("aTitle") & "]]></Caption>" _
			& "<xImgFile>" & RSlist("xImgFile") & "</xImgFile></Reference>"

		RSlist.moveNext
	wend
	response.write "</ReferenceList>"
  end if
	
  keywordRelation RS("iCUItem")
  if not RSKey.eof then
	response.write "<RelatedList>" & vbCRLF
	while not RSKey.eof
		response.write "<RelatedRef><URL>content.asp?CuItem=" & RSKey("iCUItem") & "&amp;mp=" & request("mp") _
			& "</URL><keyWeights>" & RSKey("KeyWeights") _
			& "</keyWeights><cPostDate>" & d7Date(RSKey("xPostDate")) _
			& "</cPostDate><Caption><![CDATA[" & RSKey("sTitle") & "]]></Caption></RelatedRef>"

		RSKey.moveNext
	wend
	response.write "</RelatedList>"
  end if


  for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
	processXDataSet
  next

function keywordRelation(iCUItem)
    '----參數 CuDTGenric.iCUItem
    '----傳回值 true:產生RSKey recordsets(10筆);false:不能產生RSKey recordsets
    '----RSKey欄位
    '--------iCUItem	相關資料自動編號ID值
    '--------sTitle	相關資料標題
    '--------KeyCounts	相關資料之關鍵字詞與搜尋關鍵字詞相同之個數
    '--------KeyWeights	相關資料之關鍵字詞與搜尋關鍵字詞相同之權重總計
    xiCUItem=iCUItem
    SQLK="Select xKeyword, refID from CuDtGeneric where iCUItem =" & xiCUItem
    set RSK=conn.execute(SQLK)
    if not isNull(RSK("xKeyword")) and RSK("xKeyword") <> "" then
    xRefID = RSK("refID")
    if xRefID="" OR isNull(xRefID)	then	xRefID=xiCUItem
	'----組合where子句所需xKeyword字串
	xKeywordStr = ""
	xKeywordArray = split(RSK("xKeyword"),",")
	for i = 0 to ubound(xKeywordArray)
		'----去除權重符號
		xPos = instr(xKeywordArray(i),"*")
		if xPos <> 0 then
			xStr = Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr = trim(xKeywordArray(i))
		end if
		xKeywordStr = xKeywordStr + "'" + xStr + "',"
	next
	xKeywordStr = Left(xKeywordStr,Len(xKeywordStr)-1)
	'----產生RSKey recordsets
	SQLKey = "SELECT Top 10 CDTK.iCUItem,CDTG.sTitle,max(CDTG.xPostDate) AS xPostDate,count( CDTK.iCUItem) KeyCounts,sum(weight) KeyWeights " & _
		"FROM CuDTKeyword CDTK join CuDtGeneric CDTG On CDTK.iCUItem=CDTG.iCUItem " & _
		"where CDTK.iCUItem <> " & cStr(xiCUItem) & _
		" and CDTK.iCUItem <> " & xRefID & _
		" and CDTK.xKeyword in (" & xKeywordStr & ") " & _
		" and (CDTG.refID is NULL OR CDTG.refID <> " & cStr(xiCUItem) & ") " & _
		" Group by CDTK.iCUItem,CDTG.sTitle Order by KeyWeights DESC ,KeyCounts DESC, CDTK.iCUItem"	
	set RSKey = conn.execute(SQLKey)
	keywordRelation = true
    else
    set RSKey = conn.execute("SELECT * FROM AP WHERE 1=2")
	keywordRelation = false    
    end if
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

sub doCP
	sql = "SELECT  htx.*, xr1.deptName, u.CtUnitName " _
		& ", (SELECT count(*) FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y' AND dhtx.xiCUItem=htx.iCuItem) AS attachCount " _
		& ", (SELECT count(*) FROM CuDTPage AS phtx" _
		& " WHERE bList='Y' AND phtx.xiCUItem=htx.iCuItem) AS pageCount " _
		& " FROM CuDtGeneric AS htx LEFT JOIN dept AS xr1 ON xr1.deptID=htx.idept" _
		& " LEFT JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " WHERE htx.iCuItem=" & pkstr(xItem,"")
'	response.write sql
	set RS = conn.execute(sql)
	if not RS.eof then
	cPostDate = d7date(RS("xPostDate"))
	scPostDate = "中華民國 " & mid(cPostDate,1,2) & " 年 " & mid(cPostDate,4,2) & " 月 " & mid(cPostDate,7,2) & " 日" 
%>
		<MainArticle iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">
			<Caption><%=RS("sTitle")%></Caption>
			<Content><![CDATA[<%=ReplaceAndFindKeyword(message(RS("xBody")))%>]]></Content>
			<abstract><%=RS("xabstract")%></abstract>
			<PostDate><%=RS("xPostDate")%></PostDate>
			<cPostDate><%=scPostDate%></cPostDate>
 			<DeptName><%=RS("deptName")%></DeptName>
			<TopCat><%=RS("TopCat")%></TopCat>
			<vgroup><%=RS("vgroup")%></vgroup>
			<CtUnitName><%=deAmp(RS("CtUnitName"))%></CtUnitName>
			<xURL><%=deAmp(RS("xURL"))%></xURL>
			<xKeyword><%=deAmp(RS("xKeyword"))%></xKeyword>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile>public/Data/<%=RS("xImgFile")%></xImgFile>
<%		end if %>
		</MainArticle>
<%		
	end if
end sub

sub doFP

	xSelect = "D.deptName,htx.*, ghtx.*"
	xFrom = nullText(DSDrefModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
			& " LEFT JOIN Dept AS D ON D.deptId=ghtx.idept "
	xrCount = 0
	for each param in DSDrefModel.selectNodes("fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId='" & param.selectSingleNode("refLookup").text & "'"
		'response.write SQL  & "<HR>"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("codeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("codeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("codeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("codeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("codeSrcFld") & "='" & RSLK("codeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	for each param in DSDallModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId='" & param.selectSingleNode("refLookup").text & "'"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("codeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("codeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("codeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("codeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("codeSrcFld") & "='" & RSLK("codeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	fSql = "SELECT (SELECT count(*) FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y' AND dhtx.xiCUItem=ghtx.iCuItem) AS attachCount " _
		& ", (SELECT count(*) FROM CuDTPage AS phtx" _
		& " WHERE bList='Y' AND phtx.xiCUItem=ghtx.iCuItem) AS pageCount, " 
	fSql = fSql & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE ghtx.iCuItem=" & pkstr(xItem,"")
	fSql = fSql & " ORDER BY xPostDate DESC"

'response.write fSql & "<HR>"
'response.end
	set RS = conn.execute(fSql)

		xInDateRange = "Y"
'		if not isNull(RS("m011_edate")) AND RS("m011_edate")<>"" _
'			AND (xStdDay(RS("m011_edate")) < xStdDay(date())) then	xInDateRange="N"
		if not isNull(RS("xPostDateEnd")) then
			if RS("xPostDateEnd") <> "" AND (xStdDay(RS("xPostDateEnd")) < xStdDay(date())) then	xInDateRange="N"
		end if

	
	if not RS.eof then
	xrCount = 0
%>
		<MainArticle iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>" xInDateRange="<%=xInDateRange%>">
<%	for each param in nxml0.selectNodes("//fieldList/field")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
			xrCount = xrCount + 1
			kf = "xref"  & kf
		elseif nullText(param.selectSingleNode("fieldName"))="idept" then
			kf="deptName"
		end if	
%>                  		
               <MainArticleField>	
				<fieldName><%=param.selectSingleNode("fieldName").text%></fieldName>
				<Title><%=param.selectSingleNode("fieldLabel").text%></Title>
				<%'Bob 只有內文才顯示小百科 %>
				<%if param.selectSingleNode("fieldName").text = "xbody" then %>
				<Value><![CDATA[<%=ReplaceAndFindKeyword( message(RS(kf)))%>]]></Value>
				<%else%>
				<Value><![CDATA[<%=message(RS(kf))%>]]></Value>
				<%end if%>
               </MainArticleField>				
            	
<%
	next
%>		
		</MainArticle>
		<!--發佈主題館-->
<%
	sqlsubject="SELECT DISTINCT CatTreeRoot.CtRootName, CatTreeRoot.CtRootID FROM KnowledgeToSubject LEFT OUTER JOIN CatTreeRoot ON KnowledgeToSubject.subjectId = CatTreeRoot..CtRootID "
	sqlsubject=sqlsubject&" WHERE (KnowledgeToSubject.iCUItem = "& pkstr(xItem,"") &")"
	set rs_subject=conn.execute(sqlsubject)
	response.write "<SubjectList>"
	while not rs_subject.eof 
%>
			<Subject>
				<subjectname><![CDATA[<%=rs_subject("CtRootName")%>]]></subjectname>
				<subjectmp><![CDATA[<%=rs_subject("CtRootID")%>]]></subjectmp>
			</Subject>	
	<%
		rs_subject.movenext
	wend
	%>
		</SubjectList>
			<!--發佈主題館end-->
<%		
	end if
end sub

'Added By Leo  2011-07-13       回列表頁、往上一篇、往下一篇    Start
sub setPageURL_CuDTx7

    'Added By Leo
    if request.QueryString("topCat") <> "" then
        session("temptopCat") =  request.QueryString("topCat")
    else
        session("temptopCat") =  ""
    end if
    strQueryiCtUnit = "SELECT CtUnitID FROM CatTreeNode WHERE ctnodeid = " & Request.QueryString("ctNode")
    Set iCtUnitRead = conn.execute(strQueryiCtUnit)
    if not iCtUnitRead.eof then
        strCreateScript = "SELECT htx.* , ghtx.* , u.CtUnitName , xr1.deptName , " _
                        & "xref1.mValue AS xreffctupublic , xref2.mValue AS xrefNewtopCat " _
                        & "into #tmpCuDtGeneric FROM ((CuDTx7 AS htx " _
                        & "JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
                        & "JOIN CtUnit AS u ON u.CtUnitID=ghtx.iCtUnit " _
                        & "LEFT JOIN dept AS xr1 ON xr1.deptID=ghtx.idept "_
                        & "LEFT JOIN codeMain AS xref1 ON xref1.mCode = ghtx.fctupublic AND xref1.codeMetaID='isPublic') " _
                        & "LEFT JOIN CodeMain AS xref2 ON xref2.mCode = ghtx.topCat AND xref2.codeMetaID='new_life') "_
                        & "	WHERE ghtx.fCTUPublic='Y' AND iCtUnit = " & iCtUnitRead("CtUnitID")
        if session("temptopCat") <> "" then
            strCreateScript = strCreateScript & " AND NewtopCat LIKE '%" & session("temptopCat") & "%' ORDER BY xPostDate DESC "
        else
            strCreateScript = strCreateScript & " ORDER BY xPostDate DESC "
        end if
        
        strDropScript = " DROP TABLE #tmpCuDtGeneric"
        strBackScript = strCreateScript & " select top 1 * from #tmpCuDtGeneric where giCuItem > " & Request.QueryString("xItem") & " and showType = 1 ORDER BY giCUItem" & strDropScript
        strNextScript = strCreateScript & " select top 1 * from #tmpCuDtGeneric where giCuItem < " & Request.QueryString("xItem") & " and showType = 1 ORDER BY giCUItem DESC" & strDropScript      
            set BackRecord = conn.execute(strBackScript)
	        if not BackRecord.eof then
                BackURL = "/ct.asp?xItem=" & BackRecord("iCuItem") & "&amp;ctNode=" & Request.QueryString("ctNode") & "&amp;mp=1&amp;kpi=0"
    %>
                <BackURL><%=BackURL %></BackURL>
    <%        
            end if

            set NextRecord = conn.execute(strNextScript)
	        if not NextRecord.eof then
                NextURL = "/ct.asp?xItem=" & NextRecord("iCuItem") & "&amp;ctNode=" & Request.QueryString("ctNode") & "&amp;mp=1&amp;kpi=0"
    %>
                <NextURL><%=NextURL %></NextURL>
    <%        
            end if
    END if

end sub

sub setPageURL_CuDTxT

    'Added By Leo
    if request.QueryString("topCat") <> "" then
        session("temptopCat") =  request.QueryString("topCat")
    else
        session("temptopCat") =  ""
    end if
    strQueryiCtUnit = "SELECT CtUnitID FROM CatTreeNode WHERE ctnodeid = " & Request.QueryString("ctNode")
    Set iCtUnitRead = conn.execute(strQueryiCtUnit)
    if not iCtUnitRead.eof then
    if Request.QueryString("ctNode") = 1580 then
        strCreateScript = "SELECT htx.* , ghtx.* , u.CtUnitName , xr1.deptName , " _
                        & "xref1.mValue AS xreffctupublic , xref2.mValue AS xrefNewtopCat " _
                        & "into #tmpCuDtGeneric FROM ((CuDTx5 AS htx " _
                        & "JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
                        & "JOIN CtUnit AS u ON u.CtUnitID=ghtx.iCtUnit " _
                        & "LEFT JOIN dept AS xr1 ON xr1.deptID=ghtx.idept "_
                        & "LEFT JOIN codeMain AS xref1 ON xref1.mCode = ghtx.fctupublic AND xref1.codeMetaID='isPublic') " _
                        & "LEFT JOIN CodeMain AS xref2 ON xref2.mCode = ghtx.topCat AND xref2.codeMetaID='new_life') "_
                        & "	WHERE ghtx.fCTUPublic='Y' AND iCtUnit = " & iCtUnitRead("CtUnitID")
    else
        strCreateScript = "SELECT htx.* , ghtx.* , u.CtUnitName , xr1.deptName , " _
                        & "xref1.mValue AS xreffctupublic , xref2.mValue AS xrefNewtopCat " _
                        & "into #tmpCuDtGeneric FROM ((CuDTxT AS htx " _
                        & "JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
                        & "JOIN CtUnit AS u ON u.CtUnitID=ghtx.iCtUnit " _
                        & "LEFT JOIN dept AS xr1 ON xr1.deptID=ghtx.idept "_
                        & "LEFT JOIN codeMain AS xref1 ON xref1.mCode = ghtx.fctupublic AND xref1.codeMetaID='isPublic') " _
                        & "LEFT JOIN CodeMain AS xref2 ON xref2.mCode = ghtx.topCat AND xref2.codeMetaID='new_life') "_
                        & "	WHERE ghtx.fCTUPublic='Y' AND iCtUnit = " & iCtUnitRead("CtUnitID")
    end if
        if session("temptopCat") <> "" then
            strCreateScript = strCreateScript & " AND NewtopCat LIKE '%" & session("temptopCat") & "%' ORDER BY xPostDate DESC "
        else
            strCreateScript = strCreateScript & " ORDER BY xPostDate DESC "
        end if
        strDropScript = " DROP TABLE #tmpCuDtGeneric"
        strBackScript = strCreateScript & " select top 1 * from #tmpCuDtGeneric where giCuItem > " & Request.QueryString("xItem") & " and showType = 1 ORDER BY giCUItem" & strDropScript
        strNextScript = strCreateScript & " select top 1 * from #tmpCuDtGeneric where giCuItem < " & Request.QueryString("xItem") & " and showType = 1 ORDER BY giCUItem DESC" & strDropScript

           set BackRecord = conn.execute(strBackScript)
	        if not BackRecord.eof then
                BackURL = "/ct.asp?xItem=" & BackRecord("iCuItem") & "&amp;ctNode=" & Request.QueryString("ctNode") & "&amp;mp=1&amp;kpi=0"
    %>
                <BackURL><%=BackURL %></BackURL>
    <%        
            end if

            set NextRecord = conn.execute(strNextScript)
	        if not NextRecord.eof then
                NextURL = "/ct.asp?xItem=" & NextRecord("iCuItem") & "&amp;ctNode=" & Request.QueryString("ctNode") & "&amp;mp=1&amp;kpi=0"
    %>
                <NextURL><%=NextURL %></NextURL>                
    <%        
            end if
    END if

end sub
'Added By Leo  2011-07-13       回列表頁、往上一篇、往下一篇     End

sub dorefDataGroup(rNode)

	response.write "<refDataBlock>"
	response.write rNode.seleitSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml

	xSelect = nullText(rNode.selectSingleNode("sqlSelect"))
	xFrom = nullText(rNode.selectSingleNode("sqlFrom"))
	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	xWhere = " WHERE 1=1"
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		xWhere = xWhere & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next
	
	fSql = "SELECT " & HeaderCount & xSelect & " FROM " & xFrom & xWhere
		
	
	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"


 Set RSreg = conn.execute(fSql)

		for xi = 0 to RSreg.fields.count-1
			if nullText(rNode.selectSingleNode("groupField[text()='"&RSreg.fields(xi).name&"']"))="" then
				response.write "<nF>" & RSreg.fields(xi).name & "</nF>"
			end if
		next
	orgCPKey = ""
	while not RSreg.eof
		cpKey = ""
		for each xf in rNode.selectNodes("groupKey")
			cpKey = cpKey & RSreg(nullText(xf))
		next
		if cpKey <> orgCPKey then
			if orgCPKey <>"" then	response.write "</refGroup>"
			response.write "<refGroup>"
			for each xf in rNode.selectNodes("groupField")
%>                  
            		<groupField>		
				<fieldName><%=nullText(xf)%></fieldName>
				<Value><![CDATA[<%= message(RSreg(nullText(xf))) %>]]></Value>
            		</groupField>
<%  		next
			orgCPKey = cpKey
		end if
%>
		<refData>
<%
		for xi = 0 to RSreg.fields.count-1
			if nullText(rNode.selectSingleNode("groupField[text()='"&RSreg.fields(xi).name&"']"))="" then
%>                  
            		<ArticleField>		
				<fieldName><%=RSreg.fields(xi).name%></fieldName>
				<Value><![CDATA[<%= RSreg.fields(xi) %>]]></Value>
            		</ArticleField>
    <%  	end if 
    	next%>
 		</refData>   
<%
         RSreg.moveNext
	wend

	if orgCPKey <>"" then	response.write "</refGroup>"
	response.write "</refDataBlock>"
end sub

sub dorefDataBlock(rNode)

	response.write "<refDataBlock>"
	response.write rNode.selectSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml

	xSelect = nullText(rNode.selectSingleNode("sqlSelect"))
	xFrom = nullText(rNode.selectSingleNode("sqlFrom"))
	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	xWhere = " WHERE 1=1"
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		xWhere = xWhere & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next
	
	fSql = "SELECT " & HeaderCount & xSelect & " FROM " & xFrom & xWhere
		
	
	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"

PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=60  
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
	RSreg.CacheSize = PerPageSize
RSreg.Open fSql,Conn,3,1

if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
end if   

%>
        <qURL>xItem=<%=xItem%>&amp;ctNode=<%=ctNode%></qURL> 
	<nowPage><%=nowPage%></nowPage>
	<totPage><%=totPage%></totPage>
	<totRec><%=totRec%></totRec>
	<PerPageSize><%=PerPageSize%></PerPageSize>
<%		
	
'	set RSreg = conn.execute(fSql)
If not RSreg.eof then   

    for i=1 to PerPageSize
%>
		<refData>
<%
		for xi = 0 to RSreg.fields.count-1
%>                  
            		<ArticleField>		
				<fieldName><%=RSreg.fields(xi).name%></fieldName>
				<Value><![CDATA[<%= RSreg.fields(xi) %>]]></Value>
            		</ArticleField>
    <%  next%>
 		</refData>   
<%
         RSreg.moveNext
         if RSreg.eof then exit for
	next 
end if


	response.write "</refDataBlock>"
end sub

function nullAttribute(xnode, xname)
on error resume next
	xs = ""
	xs = xnode.getAttributeNode(xname).value
	nullAttribute = xs
end function

sub xdorefDataBlock(rNode)

	response.write "<refDataBlock>"
	response.write rNode.selectSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml
	xSelect = ""
	xFrom = ""
	for each xj in rNode.selectNodes("Entity")
		tablename = nullText(xj.selectSingleNode("table"))
		xFrom = xFrom & " " & nullText(xj.selectSingleNode("join")) & " " & tableName
		myAlias = nullText(xj.selectSingleNode("alias"))
		if myAlias="" then	
			myAlias = tablename
		else
			xFrom = xFrom & " AS " & myAlias
		end if
		myAlias = myAlias & "."
		if nullText(xj.selectSingleNode("join"))<>"" then
		end if
		
		for each pkf in xj.selectNodes("pickField")
			myField = nullAttribute(pkf, "orgField") 
			if myField="" then
				xSelect = xSelect & "," & myAlias & nullText(pkf)
			else
				xSelect = xSelect & "," & myAlias & myField & " AS " & nullText(pkf)
			end if
		next
	next

	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	fSql = "SELECT " & HeaderCount & mid(xSelect,2) & " FROM " & xFrom 
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"
	response.write "</refDataBlock>"
end sub

sub xxxx
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		fSql = fSql & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next

	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"
	
PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=60  
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
	RSreg.CacheSize = PerPageSize
RSreg.Open fSql,Conn,3,1

if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
end if   
	


%>
        <qURL>xItem=<%=xItem%>&amp;ctNode=<%=ctNode%></qURL> 
	<nowPage><%=nowPage%></nowPage>
	<totPage><%=totPage%></totPage>
	<totRec><%=totRec%></totRec>
	<PerPageSize><%=PerPageSize%></PerPageSize>
<%		
	
'	set RSreg = conn.execute(fSql)
If not RSreg.eof then   

    for i=1 to PerPageSize

    	xURL = "ct.jsp?xItem=" & RSreg("iCuItem") & "&amp;ctNode=" & nullText(rNode.selectSingleNode("DataNode"))
     	if lcase(xBaseTableName) = "adrotate" then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("ibaseDSD") = 2 then	xURL = deAmp(RSreg("xURL"))
	   	if RSreg("showType") = 2 then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("showType") = 3 then	xURL = "public/Data/" & RSreg("fileDownLoad")
%>
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>">
			<xURL newWindow="<%=RSreg("xNewWindow")%>"><%=xURL%></xURL>
<%
		xrCount = 0
		for each param in rbDSDDom.selectNodes("//field[showListClient!='']") 
			kf = nullText(param.selectSingleNode("fieldName"))
			if nullText(param.selectSingleNode("refLookup")) <> "" _
				AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
				AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
				xrCount = xrCount + 1
				kf = "xref" & xrCount & kf
			end if	    	
%>                  
            		<ArticleField>		
				<fieldName><%=kf%></fieldName>
<%
    kfvalue=RSreg(kf) %>
				<Value><![CDATA[<%= kfvalue %>]]></Value>
            		</ArticleField>
    <%  next%>
 		</Article>   
<%
         RSreg.moveNext
         if RSreg.eof then exit for
	next 
end if
end sub
qStr = request.QueryString
if Instr(qStr, "&memID") > 0 then
	qStr = mid(qStr, 1, Instr(qStr, "&memID") - 1 )
end if
response.write "<qStr>" & deAmp(qStr) & "</qStr>"
%>
<!--#include file="x1Menus.inc" -->
<%	
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count = rs("counts") + 1
		'sql = "UPDATE counter SET counts = counts + 1  WHERE mp='" & request("mp") & "'"
	Else
		count = 1
		'把count移到viewcounter.aspx 
		'sql="INSERT INTO counter (mp, counts) VALUES ('" & request("mp") & "','1')"
	End If
	'Set rs = conn.Execute(sql)
	sql = " Select SUM(最新消息)+SUM(農業與生活)+SUM(新優質農家)+SUM(產銷專欄)+SUM(資源推薦)+ "
	sql = sql + "SUM(影音專區)+SUM(相關網站)+SUM(農業小百科)+SUM(農業知識家)+SUM(農業知識庫) +SUM(農作物地圖) + SUM(主題館) AS Total from CounterByDate "
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count2 = CLng(rs("Total"))
	Else
		count2 = 1
	End If
	count = CLng(count) + CLng(count2)
	Response.Write "<Counter>" & count & "</Counter>"
	
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	sql = "SELECT max(xpostDate) FROM CuDtGeneric AS htx JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit" _
		& " AND n.CtRootID=" & xRootID
	set rs = conn.execute(sql)
%>
<lastupdate><% =year(rs(0)) & "/" & month(rs(0)) &"/" & day(rs(0)) %></lastupdate>
<etoday><%=date()%></etoday>
<today><% ="民國"& year(date())-1911 & "年" & month(date()) &"月"& day(date()) & "日"%></today>
</hpMain>
