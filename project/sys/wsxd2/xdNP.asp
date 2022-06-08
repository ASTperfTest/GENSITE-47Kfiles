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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdNP.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("ctNode", "mp", "CtNode")
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
<!--#Include virtual = "/inc/MSclient.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

npURL = ""

sql = "SELECT n.CtUnitID, n.CtNodeKind, n.CtNodeNPKind, u.iBaseDSD, u.CtUnitKind" _
	& " FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON n.CtUnitID=u.CtUnitID" _
	& " WHERE n.CtNodeID=" & pkStr(request("ctNode"),"")
set RS = conn.execute(sql)
if RS("CtNodeKind") = "U" then
	if RS("CtUnitKind") = "1" then
		sql = "SELECT TOP 1 * FROM CuDTGeneric WHERE iCtUnit=" & RS("CtUnitID") _
			& " ORDER BY xImportant DESC"
		set RSx = conn.execute(sql)
		if not RSx.eof then
			npURL = "ct.asp?xItem=" & RSx("iCuItem") & "&CtNode=" & request("ctNode")& "&mp=" & request("mp")
		end if
	else
			npURL = "lp.asp?CtNode=" & request("CtNode") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&mp=" & request("mp")
	end if
	response.write "<npURL>"&deAmp(npURL)&"</npURL>"
elseif IsNull(RS("CtNodeNPKind")) then
	sql = "SELECT * FROM CatTreeNode WHERE DataParent=" & pkStr(request("ctNode"),"") _
		& " ORDER BY CatShowOrder"
	set RS = conn.execute(sql)
	npURL = "np.asp?ctNode=" & RS("CtNodeID")& "&mp=" & request("mp")
	response.write "<npURL>"&deAmp(npURL)&"</npURL>"
else

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
	myTreeNode = request("ctNode")
%>
	<xslpage>np</xslpage>
	<%

	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(request("ctNode"),"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
'	xParent = RS("CtNodeID")
	myParent = xParent
	xLevel = RS("DataLevel") - 1
	if RS("CtNodeKind") <> "C" then
		xLevel = xLevel -1
		myTreeNode = xParent
	end if
	upParent = 0
	myupParent = RS("CtNodeID")
	
	while xParent<>0
		sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		if RS("DataLevel") = xLevel then	upParent = xParent
		xPathStr = "<xPathNode Title=""" & deAmp(RS("CatName")) & """ xNode=""" & RS("ctNodeID") & """ />" & xPathStr
		xParent = RS("DataParent")
	wend
	response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>"
  
	for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
		processXDataSet
	next
%>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" myupParent="<%=myupParent%>"/>
	<MenuBar3 myTreeNode="<%=myTreeNode%>" Caption="<%=deAmp(xCtUnitName)%>">
		<%
	if isNull(xRootID) then _
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataParent=" & myupParent _
		& " AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	while not RS.eof
			if RS("CtUnitKind") = "K" then
		   		xURL = "kmDoit.asp?"&RS("redirectURL")
		   	elseif RS("CtNodeKind") = "C" then		'-- Folder
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&mp=" & request("mp")
			else
				if RS("redirectURL")<> "" then
					xUrl = RS("redirectURL")
				elseif RS("CtUnitKind") ="2" then
					xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
				elseif isNumeric(RS("iBaseDSD")) then
					xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
				end if
	   	   end if
%>
		<MenuCat newWindow="<%=RS("newWindow")%>">
			<Caption>
				<%=deAmp(RS("CatName"))%>
			</Caption>
			<redirectURL>
				<%=deAmp(xUrl)%>
			</redirectURL>
			<%
			if RS("CtNodeID") = cint(myTreeNode)  then
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.inUse='Y' AND a.DataParent=" & myTreeNode & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&mp=" & request("mp")
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")& "&amp;mp=" & request("mp")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")& "&amp;mp=" & request("mp")
							end if
				   	   end if
%>
			<MenuItem newWindow="<%=RSx("newWindow")%>">
				<Caption>
					<%=deAmp(RSx("CatName"))%>
				</Caption>
				<redirectURL>
					<%=deAmp(xURL)%>
				</redirectURL>
			</MenuItem>
			<%						RSx.moveNext
				wend
			end if
%>
		</MenuCat>
		<%		
		RS.moveNext
	wend
%>
	</MenuBar3>
	<!--#include file="x1Menus.inc" -->
<%end if%>
</hpMain>
