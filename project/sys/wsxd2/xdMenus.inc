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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdMenus.inc"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp")
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
%>	<MenuBar myTreeNode="<%=myTreeNode%>">
<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	  SqlCom = "SELECT b.CatName AS xCat, b.ctNodeID AS xNodeID, a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataLevel=2 AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by b.CatShowOrder, a.CatShowOrder"

	response.write"<sql>"& SqlCom &"</sql>"
	set RS = conn.execute(SqlCom)
	tabindex = 0
	divindex = 0
	styletop = 235
	xCat = ""
	while not RS.eof
	        tabindex = tabindex + 1
	        divindex = divindex + 1	    
	        if RS("CtUnitKind") = "K" then
			xURL = "kmDoit.asp?"&RS("redirectURL")
		elseif RS("CtNodeKind") = "C" then		'-- Folder
			xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;mp=" & request("mp")
		else
			if RS("redirectURL")<> "" then
				xUrl = RS("redirectURL")
			elseif RS("CtUnitKind") ="2" then
				xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
			elseif isNumeric(RS("iBaseDSD")) then
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
			end if
	   	end if
		if RS("xCat") <> xCat then 
			if xCat<>"" then
				styletop = styletop + 28
%>	
		</MenuCat> 
<%			end if
			xCat = RS("xCat")			
%>
		<MenuCat>
			<Caption><%=xCat%></Caption>
			<CtNode><%=deAmp(RS("xNodeID"))%></CtNode>
			<redirectURL><%=deAmp(xUrl)%></redirectURL>
			<tabindex><%=tabindex%></tabindex>
			<divindex>div<%=divindex%></divindex>
			<Mover>MM_showHideLayers('div<%=divindex%>','','show')</Mover>
			<Mout>MM_showHideLayers('div<%=divindex%>','','hide')</Mout>
			<styletop><%=styletop%></styletop>
<%		end if 	              			
%>
				<MenuItem newWindow="<%=RS("newWindow")%>">
					<Caption><%=deAmp(RS("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
<%
		 if RS("DataParent") = cint(myTreeNode)  then
		 	styletop = styletop + 21          	
	         end if   
%>		
				</MenuItem>
<%		
		RS.moveNext
	wend
	if xCat<>"" then	response.write "</MenuCat>"
%>
	</MenuBar>
<%
	xRootID2=nullText(refModel.selectSingleNode("MenuTree2"))
	if xRootID2<>"" then

	  response.write "<MenuBar2>"	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataLevel=1 AND a.CtRootID = " & PkStr(xRootID2,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	while not RS.eof
			if RS("CtUnitKind") = "K" then
		   		xURL = "kmDoit.asp?"&RS("redirectURL")
		   	elseif RS("CtNodeKind") = "C" then		'-- Folder
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;mp=" & request("mp")
			else
				if RS("redirectURL")<> "" then
					xUrl = RS("redirectURL")
				elseif RS("CtUnitKind") ="2" then
					xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
				elseif isNumeric(RS("iBaseDSD")) then
					xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
				end if
	   	   end if
%>
			<MenuCat newWindow="<%=RS("newWindow")%>">
				<Caption><%=deAmp(RS("CatName"))%></Caption>
				<redirectURL><%=xUrl%></redirectURL>
			</MenuCat>
<%		
		RS.moveNext
	wend
%>
	</MenuBar2>
<%	end if %>
