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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\attributeVote.inc"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("te")
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
%><%
	'---Aק벼---
	response.write "<queryItems>"
	for each item in request.querystring
		response.write "<" & item & ">" & request.querystring(item) & "</" & item & ">" 
	next
	response.write "</queryItems>"
	
	response.write "<attributeVote>"
		Dim voteIctUnit : voteIctUnit = 2180
		Dim totalVote : totalVote = 0
		Dim voteA : voteA = 0
		Dim voteB : voteB = 0
		Dim voteC : voteC = 0
		Dim voteD : voteD = 0
		Dim voteE : voteE = 0
		sql = "SELECT COUNT(*) AS voteCount, topCat FROM CuDTGeneric WHERE iCtUnit = " & voteIctUnit & " " & _
					"AND refID = " & xItem & " AND fCTUPublic = 'Y' GROUP BY topCat"		
					
		set rs = conn.execute(sql)
		while not rs.eof 
			select case rs("topCat")
				case "A"
					voteA = rs("voteCount")
				case "B"
					voteB = rs("voteCount")
				case "C"
					voteC = rs("voteCount")
				case "D"
					voteD = rs("voteCount")
				case "E"
					voteE = rs("voteCount")
			end select		
			rs.MoveNext
		wend
		rs.close
		set rs = nothing
		totalVote = voteA + voteB + voteC + voteD + voteE
		response.write "<totalVote>" & totalVote & "</totalVote>"
		response.write "<voteA>" & voteA & "</voteA>"
		response.write "<voteB>" & voteB & "</voteB>"
		response.write "<voteC>" & voteC & "</voteC>"
		response.write "<voteD>" & voteD & "</voteD>"
		response.write "<voteE>" & voteE & "</voteE>"
		if totalVote > 0 then		
			response.write "<voteAPercent>" & Round(voteA / totalVote * 100, 2) & "</voteAPercent>"
			response.write "<voteBPercent>" & Round(voteB / totalVote * 100, 2) & "</voteBPercent>"
			response.write "<voteCPercent>" & Round(voteC / totalVote * 100, 2) & "</voteCPercent>"
			response.write "<voteDPercent>" & Round(voteD / totalVote * 100, 2) & "</voteDPercent>"
			response.write "<voteEPercent>" & Round(voteE / totalVote * 100, 2) & "</voteEPercent>"					
			
			'-------態度投票訊息顯示筆數文章id 281914
			sql = "select stitle from CuDTGeneric where icuitem = 281914" 
			set rs = conn.execute(sql)
			maxAttributeCount = 10
			if not rs.eof  then
				maxAttributeCount = rs("stitle")
			End if
			
			sql = "SELECT TOP " & maxAttributeCount & " iCuItem, xBody FROM CuDTGeneric WHERE iCtUnit = " & voteIctUnit & " " & _
						"AND refID = " & xItem & " AND CAST(xBody AS nvarchar) <> '' AND fCTUPublic = 'Y' ORDER BY Created_Date DESC"			
			set rs = conn.execute(sql)
			while not rs.eof 
				response.write "<article>"
				response.write "<icuitem>" & rs("iCuItem") & "</icuitem>"
				response.write "<xBody><![CDATA[" & rs("xBody") & "]]></xBody>"
				response.write "</article>"
				rs.MoveNext
			wend
		end if
	response.write "</attributeVote>"
	'---end of Aק벼---
%>	