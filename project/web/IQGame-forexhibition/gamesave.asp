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
progPath="D:\hyweb\GENSITE\project\web\IQGame-forexhibition\gamesave.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("ptime", "tone", "toneselect", "name", "money", "age", "sex")
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

    Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")
	ptime = request("ptime")
	
	tone  = "0"
	if request("tone") = "1" then
	    tone = request("toneselect")
	else
	    tone  = "0"
	end if
	
	name = pkStr(request("name"),"")
    bfind = "N"
	
	SQL = "select * from flashGame,CuDTGeneric where CuDTGeneric.iCuItem=flashGame.giCuItem and deditDate= (rtrim(ltrim(str(year(getdate())))) + '/' + rtrim(ltrim(str(month(getdate())))) +  '/' + rtrim(ltrim(str(day(getdate()))))) and name = '" & request("name") & "' order by rate desc" 
	   
	   'Response.Write(SQL)
	SET todaylogin=conn.execute(SQL)
	   
	nCount = 0
	   
	while not todaylogin.eof
	    nCount = nCount + 1
	    todaylogin.moveNext
	wend
	   'response.Write SQL & nCount
	if nCount > 3 then
	else	
        '	response.write pkStr(request("name"),"")
        ''  response.end
        'xSql = "INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,sTitle,xBody,xPostDate,xKeyword,iDept,iEditor,fctupublic,deditDate,xnewWindow,showType)" _
        '  	& " VALUES(38,651," & name & ",'',getdate(),'',0,'hyweb','Y',getdate(),'N',1)"
        xSql = "INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,sTitle,xBody,xPostDate,xKeyword,iDept,iEditor,fctupublic,deditDate,xnewWindow,showType)" _
    	    & " VALUES(38,651," & name & ",'',getdate(),'',0,'hyweb','Y',rtrim(ltrim(str(year(getdate())))) + '/' + rtrim(ltrim(str(month(getdate())))) +  '/' + rtrim(ltrim(str(day(getdate())))),'N',1)"
    	
        sql = "set nocount on;"&xSql&"; select @@IDENTITY as NewID"    
	    'debugPrint Sql
	    'response.write 	fileName & "<BR>"
        'response.write sql	& "<BR>"
        'response.end
        set RSx = conn.Execute(sql)
        xNewIdentity = RSx(0)    	
        money = request("money")
        rate = money - ptime
        str="insert into flashGame(giCuItem,name,age,sex,ptimemin,ptimesec,ptime,email,hakkalangtone,money,rate) values(" & xNewIdentity & "," & name  & "," & pkStr(request("age"),"") & "," & pkStr(request("sex"),"") & "," & Cint((ptime-(ptime mod 60))/60) & "," & Cint(ptime mod 60) & "," & ptime & "," & pkStr(request("email"),"") & "," & pkStr(tone,"") & "," & money & "," & rate & ")"
   
        conn.execute(str)
        'SQL="select top 20 * from flashGame order by rate desc,giCuItem desc"
        SQL="select top 1 * from flashGame where name = " & name   & " and money > " & money & ""
        'response.write SQL
        'response.end
        SET gamedata=conn.execute(SQL)
       
        while not gamedata.eof
            'response.write xNewIdentity
            'response.write gamedata("giCuItem")
            'if cstr(xNewIdentity) = cstr(gamedata("giCuItem")) then
            bfind = "Y"
            'end if
            gamedata.moveNext
        wend
  end if
  set conn = Nothing
  response.write "&IN=" & bfind
  
FUNCTION pkStr (s, endchar)
    if s="" then
	    pkStr = "null" & endchar
    else
	    pos = InStr(s, "'")
	    While pos > 0
		    s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		    pos = InStr(pos + 2, s, "'")
	    Wend
	    pkStr="'" & s & "'" & endchar
    end if
END FUNCTION
%>
