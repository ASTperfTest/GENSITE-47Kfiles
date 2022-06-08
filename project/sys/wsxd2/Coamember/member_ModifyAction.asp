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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\member_ModifyAction.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("account", "birthYear", "birthMonth", "birthday", "sex", "addr", "zip", "home_ext", "mobile", "actor", "member_org", "com_ext", "ptitle", "epapercheck")
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
%><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 
on error resume next
response.expires = 0 
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
	'改回舊的連線
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")
	'改回舊的連線end
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	Response.Buffer = true
	
	If Session("memID") <> "" Then

		if Request("account") <> "" and Request("passwd1") <> "" and Request("email") <> "" Then 

			'---vincent:判斷傳入的會員帳密是否正確---
			set rs = conn.Execute("select email from Member Where account = " & pkStr(Request("account"), "") & " AND passwd = " & pkStr(request("passwd1"), "") & " ")
		
			if rs.eof then
				'---傳入帳密錯誤,導回首頁---
				response.write "<script>alert('原密碼輸入錯誤!!');history.back();</script>"
			
			else
				'---傳入帳密正確---
				Dim RsEmail, FcEmail, RsCount
				RsEmail = Trim(rs("email"))
				FcEmail = Trim(Request("email"))
				If (RsEmail <> FcEmail) Then
					Set rs2 = conn.Execute("SELECT COUNT(email) AS count FROM Member WHERE (email = '" & FcEmail & "')")
					if not rs2.eof then
						RsCount = rs2("count")
						If (RsCount > 0) Then
							response.write "<script>alert(' E-mail 有誤或重覆(" & RsCount & ")，請重新輸入！');history.back();</script>"
						End If
					end If
					Set rs2 = Nothing
				End If
				Set rs = Nothing
				
				'---判斷是否有填寫出生年---
				if Request("birthYear") <> "" then
					Dim xbirtmm
					Dim xbirtdd
					Dim xbirthday
					if Request("birthMonth") <= 9 then 
						xbirtmm = "0" & Request("birthMonth")
					else
						xbirtmm = Request("birthMonth")
					end if
				
					if Request("birthday") <= 9 then 
						xbirtdd = "0" & Request("birthday")
					else
						xbirtdd = Request("birthday")
					end if			
					xbirthday = pkStr(Request("birthYear") & xbirtmm & xbirtdd, "")
					
					xbirthYear =Request("birthYear")
					birthdayForValidate = xbirthYear & "/" & xbirtmm & "/" & xbirtdd
										
					If IsDate(birthdayForValidate) = False Then
						response.write "<script>alert('請輸入正確出生日期!!');history.back();</script>"
						
					Else
						UpdateMemberData()'更新member資料 By Ivy
					End If
				else
					xbirthday = pkStr("", "")
				end if
			end if
		else
			response.write "<script>alert('請輸入正確資料!!');history.back();</script>"
		end if
			
	Else
		response.redirect "../mp.asp?mp=1"
	End if
	
Function UpdateMemberData()
	

	Dim account, passwd, nickname, sex, addr, zip, phone, home_ext, mobile, fax, email,sunRegion
	Dim actor, member_org, com_tel, com_ext, ptitle 
				
	account = pkStr(Request("account"), "")
	
	sql = ""
	 '---判斷使用者等級---
	if CInt(Request("level") ) >= 3 then
		'---判斷actor---
		If Request("actor") = "1" Or Request("actor") = "2" Or Request("actor") = "3" Then
			sql = GetCommonUpdatedSql() & "," & GetActorUpdatedSql() & "," & GetSpecialFieldUpdatedSql() & " WHERE account = " & account	 
		Else
			sql = GetCommonUpdatedSql() & "," & GetSpecialFieldUpdatedSql() & " WHERE account = " & account	 
		End If
	else
		If Request("actor") = "1" Or Request("actor") = "2" Or Request("actor") = "3" Then
			sql = GetCommonUpdatedSql() & "," & GetActorUpdatedSql() & " WHERE account = " & account	 
		Else
			sql = GetCommonUpdatedSql() & " WHERE account = " & account	
		End If
	
	end if

	conn.Execute(sql)
	
	'訂閱電子報處理
	epapercheck = Request("epapercheck")
	email = pkStr(Request("email"), "")
	if epapercheck = "1" then
		checksql = "select * from Epaper where email = "& email
		set check_epaper = conn.execute(checksql)
		if check_epaper.eof then
			sql1 = "INSERT INTO Epaper (email,createtime,CtRootID) VALUES ("&email&", getdate(),'21')"
			conn.execute(sql1)
		end if
	else
		sql2 = "DELETE FROM Epaper WHERE email = "& email
		conn.execute(sql2)
	end if
	'訂閱電子報處理end

	'處理 從登入失敗流程導過來的 儲存個人資料後續導頁動作
	If request("unlogin") = "true" Then
	mp =  Request("mp") 	
	directUrl = session("myURL") & "loginact.asp?redirecturl=/mp.asp?mp=" & mp & "&account2=" & Request("account") & "&passwd2=" & Request("passwd1")
	
%>
   	<script language="javascript">
			alert("會員資料填妥完畢,修改成功");
			window.location.href ="<%=directUrl%>";
   	</script>   
<%
	Else
		response.write "<script>alert('修改個人資料成功');window.location.href='" & session("myURL") & "';</script>"
	End If
	
End Function

Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function	

Function GetCommonUpdatedSql()
	
	if Request("passwd2") <> "" Then
		passwd = pkStr(Request("passwd2"), "")
	else
		passwd = pkStr(Request("passwd1"), "")
	end if
	nickname = pkStr(Request("nickname"), "")
	sex = pkStr(Request("sex"), "")
	addr = pkStr(Request("addr"), "")
	zip = pkStr(Request("zip"), "")
	phone = pkStr(Request("phone"), "")
	home_ext = pkStr(Request("home_ext"), "")
	mobile = pkStr(Request("mobile"), "")
	fax = pkStr(Request("fax"), "")
	email = pkStr(Request("email"), "")
	sunRegion = pkStr(Request("sunRegion"), "")
	Dim showCursorIcon
	If (Request("cursorcheck") <> "") Then
		showCursorIcon = 1
	Else
		showCursorIcon = 0
	End If

	GetCommonUpdatedSql = "UPDATE Member SET passwd = " & passwd & ", nickname = " & nickname & ", birthday = " & xbirthday & ", " & _
				"sex = " & sex & ", homeaddr = " & addr & ", zip = " & zip & ", phone = " & phone & ", home_ext = " & home_ext & ", " & _
				"mobile = " & mobile & ", fax = " & fax & ", email = " & email & ", keyword = " & sunRegion & ", ShowCursorIcon = " & showCursorIcon

	If (RsEmail <> Trim(Request("email"))) Then
		GetCommonUpdatedSql = GetCommonUpdatedSql & ", mcode = NULL"
	End If


End Function

Function GetActorUpdatedSql()

	actor = pkStr(Request("actor"), "")
	member_org = pkStr(Request("member_org"), "")
	com_tel = pkStr(Request("com_tel"), "")
	com_ext = pkStr(Request("com_ext"), "")
	ptitle = pkStr(Request("ptitle"), "")
			
	GetActorUpdatedSql = " actor = " & actor & ", member_org = " & member_org & ", " & _
						"com_tel = " & com_tel & ", com_ext = " & com_ext & ", ptitle = " & ptitle 
End Function

Function GetSpecialFieldUpdatedSql()
	
	introduce = pkStr(Request("introduction"), "")
	contact = pkStr(Request("contact"), "")
		
	GetSpecialFieldUpdatedSql = "introduce = " & introduce & ", contact = " & contact 
End Function
%>
