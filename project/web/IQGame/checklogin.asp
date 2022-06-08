<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/9 上午 10:55:22
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\web\IQGame\checklogin.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("kpi", "userIdno","userEmail")
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
%><%
  Response.CacheControl = "no-cache" 
  Response.AddHeader "Pragma", "no-cache" 
  Response.Expires = -1  
  %>
  <!--#Include virtual = "/inc/dbFunc.inc" -->
  <%
	if request.querystring("kpi") <> "0" then
		response.redirect "/Kpi/KpiIQGameLogin.aspx?account=" & stripHTML(request("userIdno")) & "&password=" & stripHTML(request("userEmail"))
		response.end
	else
		
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open session("ODBCDSN")
		SQL="select top 1 * from Member where account='" & stripHTML(request("userIdno")) & "' and passwd='" & stripHTML(request("userEmail")) & "'"
		SET gamedata=conn.execute(SQL)
		i = 0
		if not gamedata.eof then
			count = gamedata("logincount")+1
	   
			Today = Year(Now()) & "/" & Month(Now()) & "/" & Day(Now()) 
	   
			'SQL = "select * from MemberFlashLoginLog where account='" & request("userIdno") & "' and LoginDatetime >= '" & Today & "' and LoginDatetime < DATEADD(day, 1, '" & Today & "')"

      SQL = "select top 1 * from flashGame,CuDTGeneric where CuDTGeneric.iCuItem=flashGame.giCuItem and name = '" & stripHTML(request("userIdno")) & "' order by money desc" 
        
      SET highMoneyrs=conn.execute(SQL)
        
      highMoney = 0
      if not highMoneyrs.eof then
        highMoney = highMoneyrs("money")
      end if

	    SQL = "select * from flashGame,CuDTGeneric where CuDTGeneric.iCuItem=flashGame.giCuItem and deditDate= (rtrim(ltrim(str(year(getdate())))) + '/' + rtrim(ltrim(str(month(getdate())))) +  '/' + rtrim(ltrim(str(day(getdate()))))) and name = '" & request("userIdno") & "' order by rate desc" 
	   
			'Response.Write(SQL)
			SET todaylogin=conn.execute(SQL)
	   
			nCount = 0
	   
	    while not todaylogin.eof
	      nCount = nCount + 1
	      todaylogin.moveNext
      wend
			'response.Write SQL & nCount
			if nCount > 3 then
				Response.Write("success=true&count=" & count & "&realname=" & gamedata("realname") & "&relogin=true&highMoney=" & highMoney & "&")
			else	   
				Response.Write("success=true&count=" & count & "&realname=" & gamedata("realname") & "&relogin=false&highMoney=" & highMoney & "&")
				conn.execute("update Member set logincount=logincount+1 where  account='" & stripHTML(request("userIdno")) & "' and passwd='" & stripHTML(request("userEmail")) & "'")
				conn.execute("insert MemberFlashLoginLog(account) values('" & stripHTML(request("userIdno")) & "')")
     	end if	   
		else
			Response.Write("success=false&")
		end if  
		set conn = Nothing
	end if
%>
  
