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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\member_ApplydealForScholar.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("account", "idn", "actor", "member_org", "ptitle", "birthYear", "birthMonth", "birthday", "orderepaper", "homeaddr", "phone", "mobile", "zip", "home_ext", "sex", "com_ext", "htx_KMcat")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array("realname" ,"nickname")
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array("realname" ,"nickname")
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
<!--#include file = "des.inc" -->
<%
	'改回舊的連線
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")
	'改回舊的連線end
 	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	Response.Buffer = true
	Dim Message
	
	'檢查帳號 Grace
	IF Request("account") = "" Then
	    PrintError(1)
	ElseIf len(Request("account")) < 6 Then
	    PrintError(2)
	ElseIf len(Request("account")) > 30 Then
	    PrintError(3)
	End If
	'檢查帳號end
		
	'檢查密碼 Grace
	
	if Request("passwd") = "" then
	    PrintError(5)
	elseif len(Request("passwd")) > 16 then
	    PrintError(6)
	elseif len(Request("passwd")) < 6 then
	    PrintError(7)
	elseif len(Request("passwd2")) = "" then
	    PrintError(8)
	elseif Request("passwd2") <> Request("passwd") then
	    PrintError(9)
	else
	    dim digitflag 
		dim charflag 
		digitflag = false
		charflag = false
	    l = len(Request("passwd"))
		for i = 1 to l
		    ch = mid(Request("passwd"),i,1)	
			a = asc(ch)
		    if a = 34 or a = 32 or a = 39 or a = 38 then
			    PrintError(10)
			elseif a >= 48 and a <= 57  then '數字 
			    digitflag = true
			elseif a >= 65 and a <= 90 then '英文字母大寫
			    charflag = true
			elseif a >= 97 and a <= 122 then '英文字母小寫
				charflag = true
			end if
		next
			
	    if not digitflag then 
		    PrintError(11)
		end if
		if not charflag then 
		    PrintError(12)
		end if
	end if
	'檢查密碼end

    '用來印出檢查帳號跟密碼時的錯誤訊息 Grace
    sub PrintError(n)
	    if n = 1 then
		    response.write "<script>alert('您忘了填寫帳號了！');history.back();</script>"
		elseif n = 2 then
		    response.write "<script>alert('您所填寫的帳號少於6碼！');history.back();</script>"
		elseif n = 3 then
		    response.write "<script>alert('您所填寫的帳號超過30碼！');history.back();</script>"
		elseif n = 4 then
		    response.write "<script>alert('帳號限用英文或數字，可用『-』或『_』！');history.back();</script>"
		elseif n = 5 then
		    response.write "<script>alert('您忘了填寫密碼了！');history.back();</script>"
		elseif n = 6 then
		    response.write "<script>alert('您所填寫的密碼超過16碼！');history.back();</script>"
		elseif n = 7 then
		    response.write "<script>alert('您所填寫的密碼少於6碼！');history.back();</script>"
		elseif n = 8 then
		    response.write "<script>alert('您忘了填寫再輸入密碼了！');history.back();</script>"
		elseif n = 9 then
		    response.write "<script>alert('密碼與再輸入密碼不符！');history.back();</script>"
		elseif n = 10 then
		    response.write "<script>alert(""密碼請勿包含『\""』、『'』、『&』或空白"");history.back();</script>"
		elseif n = 11 then
		    response.write "<script>alert('密碼請至少包含一數字');history.back();</script>"
		elseif n = 12 then
		    response.write "<script>alert('密碼請至少包含一英文字');history.back();</script>"
		end if
	end sub
	
	'sam 移除idn
	if Request("account") <> "" and Request("passwd") <> "" and Request("realname") <> ""  and Request("email") <> "" and Request("actor") <> "" and Request("member_org") <> "" and Request("com_tel") <> "" and Request("ptitle") <> ""  Then
	
		email  = trim(request("email"))
        email = replace(email, "'","")
        email = replace(email, "-","")
		set rs = conn.Execute("select email from Member where email = '" & email & "'")
		If not rs.eof then
			Response.Write "<html><body bgcolor='#ffffff'>"
			Response.Write "<script language='javascript'>"
			Response.Write "alert(' E-mail 已被登記使用，請重新輸入!');"
			Response.Write "history.back();"
			Response.Write "</script>"
			Response.Write "</body></html>"   				    
		else
	
			'判斷帳號是否已使用
			set rs = conn.Execute("select * from Member Where account ='" & Request("account") & "'")
			if rs.eof then
				'判斷是否已是會員
				'Set rsid = conn.Execute("select * from Member Where id ='" & Request("idn") & "'")
				'response.Write rsid("id")
				'response.End

				if Request("birthYear") <>"" then
					Dim xbirtmm
					Dim xbirtdd
					Dim xbirthday
					if Request("birthMonth") <=9 then 
						xbirtmm ="0" & Request("birthMonth")
					else
						xbirtmm = Request("birthMonth")
					end if
					
					if Request("birthday") <=9 then 
						xbirtdd ="0" & Request("birthday")
					else
						xbirtdd = Request("birthday")
					end if
				
					xbirthday = Request("birthYear") & xbirtmm & xbirtdd
					
					'驗證日期是否正常				
					birthdayForValidate = Request("birthYear") & "/" & xbirtmm & "/" & xbirtdd
					If IsDate(birthdayForValidate) = False Then
						response.write "<script>alert('請輸入正確出生日期!!');history.back();</script>"
					Else
						validDate = true
					End If
					
				else
					xbirthday=""
				end if
				' -------2011.06.15 modify by Grace---------
				' 是否開啟動態游標
				Dim showCursorIcon 
	            If (Request("cursorcheck") <> "") Then
		        showCursorIcon = 1
	            Else
		        showCursorIcon = 0
	            End If
				if validDate = true then				
					'if rsid.eof then
						'新增帳號
						Dim orderepaper : orderepaper = "Y"
						if request("orderepaper") <> "Y" then orderepaper = "N"
						
						sql = "INSERT INTO Member(account, passwd, realname, homeaddr, phone, mobile, email, createtime, modifytime, zip, home_ext, " & _ 
									"birthday, sex, id_type1, id_type2, com_tel, com_ext, fax, create_user, KMcat, id, actor, ptitle, member_org, status, nickname, scholarValidate, orderepaper, uploadRight, uploadPicCount, keyword, ShowCursorIcon) " & _
									"VALUES( " & pkStr(Request("account"),"") & ", " & pkStr(Request("passwd"),"") & ", " & pkStr(Request("realname"),"") & ", " & pkStr(Request("homeaddr"),"") & ", " & _ 
									"" & pkStr(Request("phone"),"") & ", " & pkStr(Request("mobile"),"") & ", " & pkStr(Request("email"),"") & ", getdate(), getdate(), " & pkStr(Request("zip"),"") & ", " & _
									"" & pkStr(Request("home_ext"),"") & ", '" & xbirthday & "', " & pkStr(Request("sex"),"") & ", '1', '1', " & pkStr(Request("com_tel"),"") & ", " & _
									"" & pkStr(Request("com_ext"),"") & ", " & pkStr(Request("fax"),"") & ", 'hyweb', " & pkStr(Request("htx_KMcat"),"") & ", " & pkStr(Request("idn"),"") & ", " & _
									"" & pkStr(Request("actor"),"") & ", " & pkStr(Request("ptitle"),"") & ", " & pkStr(Request("member_org"),"") & ", 'Y', " & pkStr(Request("nickname"),"") & ", 'W', '" & orderepaper & "', 'Y', 1, " & pkStr(Request("sunRegion"),"") & "," & showCursorIcon & ")"
									
						'response.write sql
						conn.Execute(sql)

                        ' 判斷是否有 InvitationCode (invitationCode & invitationCodeType)
                        ' 新增InviteFriends_Detail
                        if Request("invitationCode") <> "" and Request("invitationCode") <> "/" and Request("invitationCodeType") <> "/" then
                            sqlInvite = ""
                            sqlInvite = sqlInvite & vbcrlf  & "declare @iCode nvarchar(50)    set @iCode=" & pkStr(Request("invitationCode"),"")
                            sqlInvite = sqlInvite & vbcrlf  & "declare @iAccount nvarchar(50) set @iAccount=" & pkStr(Request("account"),"")
                            sqlInvite = sqlInvite & vbcrlf  & "declare @iLS nvarchar(50)      set @iLS= " & pkStr(Request("invitationCodeType"),"")
                            sqlInvite = sqlInvite & vbcrlf  & ""
                            sqlInvite = sqlInvite & vbcrlf  & "INSERT INTO InviteFriends_Detail (InvitationCode, InviteAccount, LinkSource) "
                            sqlInvite = sqlInvite & vbcrlf  & "select @iCode , @iAccount , @iLS"
                            sqlInvite = sqlInvite & vbcrlf  & "where not exists (select * from InviteFriends_Detail where InvitationCode=@iCode and InviteAccount = @iAccount)"

                            conn.Execute(sqlInvite)                           
                        end if
		 
						if request("orderepaper") = "Y" then
							Dim emailFlag
							emailFlag = true
							'檢查是否已經訂閱
							set ts = conn.Execute("select count(*) from Epaper where email = '" & Request("email") & "' and CtRootID = 21")
							If Err.Number <> 0 or conn.Errors.Count <> 0 Then
								emailFlag = false
							end if		  
							if ts(0) > 0 then
								emailFlag = false
							end if
			  
							'檢查email格式是否正確
							if len(Request("email")) > 3 and InStr(Request("email"), "@") > 0 and InStr(Request("email"), ".") > 0 and emailFlag = true then
								'正確...存入資料庫
								sql = "insert into Epaper ( email, createtime, CtRootID) values ('"& Request("email") & "', getdate(), 21)"
								conn.execute(sql)	    
							end if
						end if
						
						if err.number <> 0 then
							Response.Write "<html><body bgcolor='#ffffff'>"
							Response.Write "<script language='javascript'>"
							Response.Write "alert('申請會員發生錯誤,請洽系統管理員');"
							Response.Write "location.replace('sp.asp?xdURL=coamember/member_CompleteS.asp');"			
							Response.Write "</script>"
							Response.Write "</body></html>" 
						else
						'寄出通知信  
							'response.write "<HR> 寄出認證信"
							'response.end         
							%>
							
							<!-- #Include file = "mailbody_S.inc" -->
							
							<%
							'response.end
							Response.Write "<html><body bgcolor='#ffffff'>"
							Response.Write "<script language='javascript'>"
							Response.Write "alert('申請學者會員需經管理者審核，但仍可使用您申請的帳號正常登入。\n請您至信箱收認證信，以確保您的會員權益。');"
							Response.Write "location.replace('sp.asp?xdURL=coamember/member_CompleteS.asp');"			
							Response.Write "</script>"
							Response.Write "</body></html>" 
						end if			
							
					'else '判斷是否已是會員
						'如果id已存在提示已具有會員身份
						'Response.Write "<html><body bgcolor='#ffffff'>"
						'Response.Write "<script language='javascript'>"
						'Response.Write "alert('您已具有會員身份!');"
						'Response.Write "history.back();"
						'Response.Write "location.replace('/index.jsp');"
						'Response.Write "location.replace('index.htm');"
						'Response.Write "</script>"
						'Response.Write "</body></html>"  				
					'end if
					'== 判斷是否已是會員 End ==		
				end if
					
			else
				'如果account已存在提示改用其他帳號
				Response.Write "<html><body bgcolor='#ffffff'>"
				Response.Write "<script language='javascript'>"
				Response.Write "alert('此帳號已經存在, 請改用其他帳號, 謝謝!');"
				Response.Write "history.back();"
				Response.Write "</script>"
				Response.Write "</body></html>"   
			end if 
		end if'email	
	else
		Response.Write "<html><body bgcolor='#ffffff'>"
		Response.Write "<script language='javascript'>"
		Response.Write "alert('請填入必要資料以供審核!');"
		Response.Write "history.back();"
		Response.Write "</script>"
		Response.Write "</body></html>"  
	end if 
	
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
%>
