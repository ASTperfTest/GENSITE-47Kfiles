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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\member_Modify.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("memID", "mp")
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
%><!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
	Dim SMemberId, RMemberId, Flag, SMemberPassword, RMemberPassword, myWWWsite
	
	myWWWsite = Session("myWWWsite")
	SMemberId = Session("memID")
	RMemberId = Request("memID")
	updateMemID = Request("updateMemID")
	
	Flag = true
	
	If SMemberId = "" Or RMemberId = "" Then
		Flag = false
	End If
	If SMemberId <> RMemberId Then
		Flag = false
	End If
	
	If updateMemID<>"" Then		
		Flag = true
		SMemberId = updateMemID		
	End If
	
	If Flag = false Then		
		'response.write "<script language=""javascript"" type=""text/javascript"" > alert('error');window.location.href='';</script>"
		'response.end
		'by ivy
		response.write "<script> alert('尚未登入！');window.location.href='mp.asp';</script>"
		
	End If
	'sam 因個資移除idn
	Dim account, realname, nickname, birthYear, birthMonth, birthDay, sex, addr, zip, sunRegion
	Dim phone, home_ext, mobile, fax, email, actor, member_org, com_tel, com_ext, ptitle, showCursorIcon

	if Flag = true then
		Set rsreg = conn.Execute("select * from Member Where account = '" & SMemberId & "' ")	
		
		if not rsreg.eof then
			account = trim(rsreg("account"))			
			realname = trim(rsreg("realname"))
			nickname = trim(rsreg("nickname"))
			'Dim idnLength 
			'idnLength = Len(trim(rsreg("id")))
			'if idnLength>5 then
			'	for iii = 0 to 5
			'		idn = idn & "*"
			'	next
			'end if
			'idn = idn & Mid(trim(rsreg("id")),7)
			birthYear = mid(trim(rsreg("birthday")), 1, 4)
			birthMonth = mid(trim(rsreg("birthday")), 5, 2)
			birthDay = mid(trim(rsreg("birthday")), 7, 2)
			sex = trim(rsreg("sex"))
			addr = trim(rsreg("homeaddr"))
			zip = trim(rsreg("zip"))
			phone = trim(rsreg("phone"))
			home_ext = trim(rsreg("home_ext"))
			mobile = trim(rsreg("mobile"))
			fax = trim(rsreg("fax"))
			email = trim(rsreg("email"))
			actor = trim(rsreg("actor"))
			member_org = trim(rsreg("member_org"))
			com_tel = trim(rsreg("com_tel")) 
			com_ext = trim(rsreg("com_ext"))
			ptitle =trim(rsreg("ptitle"))
			sunRegion = trim(rsreg("keyword"))
			showCursorIcon = trim(rsreg("ShowCursorIcon"))
		end if
	end If
	'檢查有沒有開啟動態游標(預設：開啟)
	Dim checkcursor
	If showCursorIcon="0" then
		checkcursor = 0
	Else
		checkcursor = 1
	End if

	'檢查有沒有訂閱電子報
	Dim checkmail
	sqlepaper = "select * from Epaper where email ='"& email &"'"
	set RSepaper = conn.execute(sqlepaper)
	if not RSepaper.eof then
		checkmail = 1
	else
		checkmail = 0
	end if 
	'檢查有沒有訂閱電子報end
	
	'=====2009/08/03 by ivy Start
	'日出日落顯示區域，從資料庫中取得
	Dim region
	sqlRegion = "select distinct isnull(xKeyword,'台北') as region from dbo.CuDTGeneric where ictunit=303"
	Set regionSorce = conn.Execute(sqlRegion)
	'=====
	
	'=====2009/08/14 by ivy 顯示會員等級及成長圖示
	Dim gradeLevel,gradeDesc,calculateTotal,picType
	
	'取得總分
	sqlCalTotal = "SELECT calculateTotal FROM MemberGradeSummary WHERE memberId = '"& account &"'"
	set RSCalTotal = conn.execute(sqlCalTotal)
	if not RSCalTotal.eof then
		calculateTotal = trim(RSCalTotal("calculateTotal"))
	else 
		calculateTotal = 0 'default
	end if
		
	
	'取得成長圖示類型
	picType = "A" 'default
	sqlStr = "SELECT pictype FROM Member WHERE account = '"& account &"'"
	set RSPicType = conn.execute(sqlStr)
	if not RSPicType.eof then
		if RSPicType("picType") <> "" then picType = trim(RSPicType("picType"))
	end if
		
	'取得等級
	sqlGrade = "SELECT TOP 1 codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'gradelevel') AND ("& calculateTotal &" >= mValue) ORDER BY mSortValue DESC"
	set RSgrade = conn.execute(sqlGrade)
	if not RSgrade.eof then	
		gradeLevel = trim(RSgrade("mCode"))
	end if
	
	'取得等級說明
	If gradeLevel = "1" Then
	  gradeDesc = "入門級會員"
	ElseIf gradeLevel = "2" Then
		gradeDesc = "進階級會員"
	ElseIf gradeLevel = "3" Then
		gradeDesc = "高手級會員"
	ElseIf gradeLevel = "4" Then
		gradeDesc = "達人級會員"
	Else
		gradeDesc = "入門級會員"
	End If
	
	'取得圖片路徑
	
	Dim picName,levelImgSrc
	If gradeLevel = 1 Then
        If picType = "A" Then
          picName = "seticon1-1.jpg"
        ElseIf picType = "B" Then
          picName = "seticon2-1.jpg"
        ElseIf picType = "C" Then
          picName = "seticon3-1.jpg"
        ElseIf picType = "D" Then
          picName = "seticon4-1.jpg"
        End If
      ElseIf gradeLevel = 2 Then
        If picType = "A" Then
          picName = "seticon1-2.jpg"
        ElseIf picType = "B" Then
          picName = "seticon2-2.jpg"
        ElseIf picType = "C" Then
          picName = "seticon3-2.jpg"
        ElseIf picType = "D" Then
          picName = "seticon4-2.jpg"
        End If
      ElseIf gradeLevel = 3 Then
        If picType = "A" Then
          picName = "seticon1-3.jpg"
        ElseIf picType = "B" Then
          picName = "seticon2-3.jpg"
        ElseIf picType = "C" Then
          picName = "seticon3-3.jpg"
        ElseIf picType = "D" Then
          picName = "seticon4-3.jpg"
        End If
      ElseIf gradeLevel = 4 Then
        If picType = "A" Then
          picName = "seticon1-4.jpg"
        ElseIf picType = "B" Then
          picName = "seticon2-4.jpg"
        ElseIf picType = "C" Then
          picName = "seticon3-4.jpg"
        ElseIf picType = "D" Then
          picName = "seticon4-4.jpg"
        End If        
      End If   
	levelImgSrc = "/xslGip/style3/images/" & picName

	'=====等級較高使用者，擁有特殊欄位編輯功能
	if CInt(gradeLevel) >= 3 then		
		Dim introduce, contact 
		personalFieldSQL = "SELECT  introduce, contact FROM Member WHERE account= '" & SMemberId & "' "
		Set rsPersonal = conn.execute(personalFieldSQL)		
		
		if not rsPersonal.eof then				
			introduce = trim(rsPersonal("introduce"))
			contact = trim(rsPersonal("contact"))
		end if
	end if
	
	'=====2009/11/20 by ivy 新增個人平均評價星等顯示
	dim personValue
	sqlPersonValue = "SELECT CAST(SUM (GradeCount / GradePersonCount) / COUNT(*) AS DECIMAL(8,2)) AVGSTART "&_
"FROM ( "&_
"SELECT CAST(KnowledgeForum.GradeCount AS DECIMAL(8,2))GradeCount, CAST(KnowledgeForum.GradePersonCount AS DECIMAL(8,2)) GradePersonCount "&_
  "FROM CuDTGeneric AS CuDTGeneric_1 "&_
             "INNER JOIN CuDTGeneric "&_
             "INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "&_
             "INNER JOIN KnowledgeForum AS KnowledgeForum_1 "&_
                                       "ON CuDTGeneric.iCUItem = KnowledgeForum_1.ParentIcuitem "&_
                                       "ON CuDTGeneric_1.iCUItem = KnowledgeForum_1.gicuitem "&_
             "INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "&_
 "WHERE (CuDTGeneric_1.iCTUnit = '933') "&_
   "AND (CuDTGeneric_1.iEditor = '" & SMemberId &"') "&_
   "AND (CodeMain.codeMetaID = 'KnowledgeType') "&_
   "AND (CuDTGeneric_1.siteId = '3') "&_
   "AND (KnowledgeForum_1.Status = 'N') "&_
   "AND (KnowledgeForum.Status = 'N') "&_
   "AND (KnowledgeForum.GradePersonCount >0) "&_
") TA  "
	
	set RSPersonValue = conn.execute(sqlPersonValue)
	if not RSPersonValue.eof then	
		personValue = trim(RSPersonValue("AVGSTART"))
	end if

	Dim personStar 
        If personValue = 0 Or Isnull(personValue) Then
			personValue = "0"
            personStar = "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue > 0 And personValue < 0.5 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_half_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue >= 0.5 And personValue <= 1 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue > 1 And personValue < 1.5 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_half_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue >= 1.5 And personValue <= 2 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue > 2 And personValue < 2.5 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_half_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue >= 2.5 And personValue <= 3 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue > 3 And personValue < 3.5 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_half_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue >= 3.5 And personValue <= 4 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf personValue > 4 And personValue < 4.5 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_half_11x11.gif"" align=""top"">"
        ElseIf personValue >= 4.5 And personValue <= 5.0 Then
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
            personStar = personStar &  "<img class=""rating"" src=""knowledge/images/icn_star_full_11x11.gif"" align=""top"">"
        End If
		
%>
<script language="JavaScript">
<!--
function send() {
	
	var form = document.Form1; 
	
	//此段拿掉的判斷,是寫錯的! realname是一個label 而非input,所以 form.realname.value是null
	//if(form.realname.value==""){
		//alert("您忘了填寫真實姓名了！"); 
		//form.realname.focus(); 
		//return false;
	//}
	//else 
	if(form.passwd1.value==""){
		alert("您忘了填寫原密碼了！"); 
		form.passwd1.focus(); 
		return false;
	}
	else if(form.passwd2.value!=""){
		
		if (form.passwd2.value.length > 16){
			alert("您所填寫的新密碼超過16碼！"); 
			form.passwd2.focus(); 
			return false;
		}
		else if(form.passwd2.value.length < 6){
			alert("您所填寫的新密碼少於6碼！"); 
			form.passwd2.focus(); 
			return false;
		}
		else if(form.passwd3.value==""){
			alert("您忘了填寫新密碼確認了！"); 
			form.passwd3.focus(); 
			return false;
		}
		else if(form.passwd2.value != form.passwd3.value){
			alert("新密碼與新密碼確認不符！");
			form.passwd3.focus(); 
			return false;
		}
		else{
			var i;
			var ch;
			var digits = "0123456789";
			var digitflag = false;
			var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			var charflag = false;
			for (i=0;i< form.passwd2.value.length;i++){		
				ch = form.passwd2.value.charAt(i);		
				if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
					alert("密碼請勿包含『\"』、『'』、『&』或空白"); 
					form.passwd2.focus();
					return false;
				}
				if( digits.indexOf(ch) >= 0 ) {
					digitflag = true;
				}					
				if( chars.indexOf(ch) >= 0 ) {
					charflag = true;
				}
			}
			if( !digitflag ) {
					alert("密碼請至少包含一數字"); 
					form.passwd2.focus();
					return false;
			}
			if( !charflag ) {
					alert("密碼請至少包含一英文字"); 
					form.passwd2.focus();
					return false;
			}
		}		
	}
	else if(form.birthYear.value==""){
			alert("您忘了填寫出生年份了！"); 
			form.birthYear.focus(); 
			return false;
	}
	else{
		
		dateObj = new Date();
		//驗證範圍
		thisYear = dateObj.getFullYear();
		
		if(form.birthYear.value.length !=4)
		{
			alert("您輸入的出生年份不足4碼！");
			form.birthYear.focus(); 
			return false;
		}
		
		if(form.birthYear.value<1850)
		{
			alert("您輸入的出生年份不正確！");
			form.birthYear.focus(); 
			return false;
		}
				
		//驗證是否包含非數字
		regObj = new RegExp("\\d{4}","g");

		matchNumber = regObj.test(form.birthYear.value)
		if(!matchNumber)
		{
			alert("您的出生年份只能輸入數字！")
			form.birthYear.focus(); 
			return false;				
		}	
		
		//驗證出生日期是否大於今天
		greaterThanToday = false;
		if(form.birthYear.value > thisYear)
		{			
			greaterThanToday = true;			
		}
		else if(form.birthYear.value == thisYear)
		{
			//取得選擇的月份						
			var objBirthMonth = document.getElementById("birthMonth")
			var birthMonth = objBirthMonth.options[objBirthMonth.selectedIndex].text;
			//取得選擇的日
			var objbirthday = document.getElementById("birthday")
			var birthDay = objbirthday.options[objbirthday.selectedIndex].text;
					
			m=dateObj.getMonth()+1;
			d=dateObj.getDate();			
					
			if(birthMonth > m)
			{
				greaterThanToday = true;
			}
			else if(birthMonth == m) 
			{
				if(birthDay > d)
				{
					greaterThanToday = true;
				}
			}
		}
		
		if(greaterThanToday == true)
		{
			alert("您輸入的出生日期不可大於今天！");
			return false;
		}
	}
	
	<% if actor = "1" or actor = "2" or actor = "3" then %>
	if( form.actor.checked == false ){
		alert("您忘了選擇身分類別了！"); 
		form.actor.focus(); 
		return false;
	}
	if(form.member_org.value==""){
		alert("您忘了填寫所屬機關名稱了！"); 
		form.member_org.focus(); 
		return false;
	}
	if(form.com_tel.value==""){
		alert("您忘了填寫所屬機關電話了！"); 
		form.com_tel.focus(); 
		return false;
	}
	if(form.ptitle.value==""){
		alert("您忘了填寫職稱了！"); 
		form.ptitle.focus(); 
		return false;
	}
	<% end if %>
	if(form.email.value==""){
		alert("您忘了填寫電子郵件了！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("@")<=-1){
		alert("您所填寫的電子郵件格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	
	if(form.contact.value != "")
	{
		var chinessRegex = /[\u4E00-\u9FA5]/g;
		if(chinessRegex.test(form.contact.value))
		{
			alert("個人/公司網站或Blog網址中不得有中文！");
			form.contact.focus();
			return false;
		}
		
		var url = /^(ftp|http|https):\/\/(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?/
		if(!url.test(form.contact.value))
		{
			alert("個人/公司網站或Blog網址格式有誤！");
			form.contact.focus();
			return false;
		}
	}
	
	return true;
}  

function BirthYearKeyPress()
{
	//是否Key in數字
	if(event.keyCode>=48 && event.keyCode<=57)
		return true;	
	else
		return false;
}
//驗證自我介紹輸入字數
function LimitTextAreaLength(field) 
{
    if (field.value.length > 1024) {
    	alert("字數不可超過1024個字");
		field.focus();
    }      
}
  
//-->
</script>

<div class="path" >目前位置：<a title="首頁" href="mp.asp">首頁</a>&gt; 修改個人資料</div>
<h3>修改個人資料</h3>
<div id="Magazine">
	<div class="Event">	
		<div class="experts">
		<form name="Form1" id = "Form1" method="post" class="FormA" action=<% if updateMemID<>"" Then %> "sp.asp?xdURL=Coamember/member_ModifyAction.asp&mp=<%=request("mp")%>&memID=<%=request("updateMemID")%>&unlogin=true"
<%Else %> "sp.asp?xdURL=Coamember/member_ModifyAction.asp&mp=<%=request("mp")%>" <%End If%> onSubmit="return send()">	
			<input type="hidden" name="account" value="<%=account%>">		
			<table  cellspacing="0" class="DataTb1" width="100%">
      <tr>
      	<th  ><label for="account">*會員帳號：</label></th>
        <td class="modifytd" ><%=account%></td>
			<td  rowspan=6 width="25%" align="left" >
				<table align="left">
				  <tr>
					<td >平均評價:</td>
					<td ><%=personStar%></td>
				  </tr>
				  <tr>
					<td ></td>
					<td >(<%=personValue%>顆星)</td>
				  </tr>
				  <tr >
					<td colspan=2><img src=<%=levelImgSrc%> alt="個人成長圖示" />
				  </td>
				  </tr>
				   <tr >
					<td align="center" colspan=2 ><%=gradeDesc%></td>
				  </tr>
				</table>
			</td>
      </tr>
      <tr>
      	<th><label for="realname">*真實姓名：</label></th>
        <td class="modifytd"><%=realname%></td>
		
      </tr>
      <!--<tr>
        <th><label for="idn">*身分證字號：</label></th>
        <td class="modifytd"></td>
		
      </tr>-->
      <tr>
        <th><label for="nickname">暱　　稱：</label></th>
        <td><input name="nickname" type="text" value="<%=nickname%>" class="Text" id="nickname" size="30"></td>
		
      </tr>
      <tr>
        <th><label for="passwd">*原密碼：</label></th>
        <td><input name="passwd1" type="password"  class="Text" id="passwd1" size="30" maxLength="16">請輸入原密碼</td>
		
      </tr>
      <tr>
        <th><label for="passwd">*設定新密碼：</label></th>
        <td >
        	<input name="passwd2" type="password" value="" class="Text" id="passwd2" size="30" maxLength="16">若不需要變更密碼，則不用填寫
        	<br/>請自訂英文（區分大小寫）、數字<br/>需同時包含至少1英文和1數字、不含空白及@，6碼以上、16碼以下。
       	</td>
      </tr>
      <tr>
        <th><label for="passwd2">*新密碼確認：</label></th>
        <td colspan=2><input name="passwd3" type="password" value="" class="Text" id="passwd3" size="30" maxLength="16">若不需要變更密碼，則不用填寫</td>
      </tr>     
      <% if actor = "1" or actor = "2" or actor = "3" then %> 
      <tr>
      	<th><label for="actor">*身分類別：</label></th>
        <td colspan=2>
        	<input name="actor" id="actor" type="radio" value="1" <% if actor = "1" then %> checked <% end if %>>研究員
      		<input name="actor" id="actor" type="radio" value="2" <% if actor = "2" then %> checked <% end if %>>教職人員
    		<input name="actor" id="actor" type="radio" value="3" <% if actor = "3" then %> checked <% end if %>>學生
    		</td>
      </tr>
      <tr>
      	<th><label for="member_org">*所屬機關名稱：</label></th>
        <td colspan=2><input name="member_org" id="member_org" type="text" size="30" class="Text" value="<%=member_org%>" ></td>
      </tr>
      <tr>
      	<th><label for="com_tel">*所屬機關電話：</label></th>
        <td colspan=2>
        	<input name="com_tel" id="com_tel" type="text" size="30" class="Text" value="<%=com_tel%>">（公）
        	<label for="com_ext">分機：</label>
					<input name="com_ext" id="com_ext" type="text" size="5" class="Text" value="<%=com_ext%>">							
				</td>
      </tr>
      <tr>
      	<th><label for="ptitle">*職稱：</label></th>
        <td colspan=2><input name="ptitle" id="ptitle" type="text" size="30"class="Text" value="<%=ptitle%>"></td>
      </tr>
      <% end if %>
	  <tr>
	  <th width="110px"><label for="sunRegion">*日出日落顯示：</label></th>
	  <td colspan=2><select name="sunRegion" id="sunRegion">
			<%do while not regionSorce.eof
				region = trim(regionSorce("region"))
			%>
			<option value="<%=region%>" <%if region=sunRegion then response.Write "selected"%>><%=region%></option>
			<% regionSorce.MoveNext
			loop
			%> 			
          </select></td>
	  </tr>
      <tr>
        <th><label for="birthday">*出生日期：</label></th>
        <td colspan=2>西元
      	  <input name="birthYear" type="text" value="<%=birthYear%>" class="Text" id="birthYear" size="5" maxLength="4" onkeypress = "return BirthYearKeyPress()">年
          <select name="birthMonth" id="birthMonth">
          	<option <% if birthMonth = "01" Then %> selected <% end if %> >1</option>
            <option <% if birthMonth = "02" Then %> selected <% end if %> >2</option>
            <option <% if birthMonth = "03" Then %> selected <% end if %> >3</option>
            <option <% if birthMonth = "04" Then %> selected <% end if %> >4</option>
            <option <% if birthMonth = "05" Then %> selected <% end if %> >5</option>
            <option <% if birthMonth = "06" Then %> selected <% end if %> >6</option>
            <option <% if birthMonth = "07" Then %> selected <% end if %> >7</option>
            <option <% if birthMonth = "08" Then %> selected <% end if %> >8</option>
            <option <% if birthMonth = "09" Then %> selected <% end if %> >9</option>
            <option <% if birthMonth = "10" Then %> selected <% end if %> >10</option>
            <option <% if birthMonth = "11" Then %> selected <% end if %> >11</option>
            <option <% if birthMonth = "12" Then %> selected <% end if %> >12</option>
          </select> 月
          <select name="birthday" id="birthday">
          	<option <% if birthday = "01" Then %> selected <% end if %> >1</option>
            <option <% if birthday = "02" Then %> selected <% end if %> >2</option>
            <option <% if birthday = "03" Then %> selected <% end if %> >3</option>
            <option <% if birthday = "04" Then %> selected <% end if %> >4</option>
            <option <% if birthday = "05" Then %> selected <% end if %> >5</option>
            <option <% if birthday = "06" Then %> selected <% end if %> >6</option>
            <option <% if birthday = "07" Then %> selected <% end if %> >7</option>
            <option <% if birthday = "08" Then %> selected <% end if %> >8</option>
            <option <% if birthday = "09" Then %> selected <% end if %> >9</option>
            <option <% if birthday = "10" Then %> selected <% end if %> >10</option>
            <option <% if birthday = "11" Then %> selected <% end if %> >11</option>
            <option <% if birthday = "12" Then %> selected <% end if %> >12</option>
            <option <% if birthday = "13" Then %> selected <% end if %> >13</option>
            <option <% if birthday = "14" Then %> selected <% end if %> >14</option>
            <option <% if birthday = "15" Then %> selected <% end if %> >15</option>
            <option <% if birthday = "16" Then %> selected <% end if %> >16</option>
            <option <% if birthday = "17" Then %> selected <% end if %> >17</option>
            <option <% if birthday = "18" Then %> selected <% end if %> >18</option>
            <option <% if birthday = "19" Then %> selected <% end if %> >19</option>
            <option <% if birthday = "20" Then %> selected <% end if %> >20</option>
            <option <% if birthday = "21" Then %> selected <% end if %> >21</option>
            <option <% if birthday = "22" Then %> selected <% end if %> >22</option>
            <option <% if birthday = "23" Then %> selected <% end if %> >23</option>
            <option <% if birthday = "24" Then %> selected <% end if %> >24</option>
            <option <% if birthday = "25" Then %> selected <% end if %> >25</option>
            <option <% if birthday = "26" Then %> selected <% end if %> >26</option>
            <option <% if birthday = "27" Then %> selected <% end if %> >27</option>
            <option <% if birthday = "28" Then %> selected <% end if %> >28</option>
            <option <% if birthday = "29" Then %> selected <% end if %> >29</option>
            <option <% if birthday = "30" Then %> selected <% end if %> >30</option>
            <option <% if birthday = "31" Then %> selected <% end if %> >31</option>
          </select> 日
				</td>
      </tr>
      <tr>
        <th><label for="sex">性別：</label></th>
        <td colspan=2>
        	<input name="sex" type="radio" value="1" id="sex"  <% if sex = "1" Then %> checked <% end if %> >男
          <input name="sex" type="radio" value="0" id="sex"  <% if sex = "0" Then %> checked <% end if %> >女
				</td>
      </tr>
      <tr>
        <th><label for="addr">地址：</label></th>
        <td colspan=2>
					<input name="addr" type="text" value="<%=addr%>" class="Text" id="addr" size="60">
          <label for="zip"> 郵遞區號：</label>
          <input name="zip" type="text" value="<%=zip%>" size="5" class="Text" id="zip">                         
				</td>
      </tr>
      <tr>
      	<th><label for="phone">電話：</label></th>
        <td colspan=2>
        	<input name="phone" type="text" value="<%=phone%>" class="Text" id="phone" size="16">
          <label for="home_ext">分機：</label>
          <input name="home_ext" type="text" value="<%=home_ext%>" class="Text" id="home_ext" size="5">範例：02-25076627
        </td>
      </tr>
      <tr>
      	<th><label for="mobile">手機號碼：</label></th>
        <td colspan=2><input name="mobile" type="text" value="<%=phone%>" class="Text" id="mobile" size="16">範例：0911123456</td>
      </tr>
      <tr>
        <th><label for="fax">傳真：</label></th>
        <td colspan=2><input name="fax" type="text" value="<%=fax%>" class="Text" id="fax" size="16">範例：02-25076627</td>
      </tr>
      <tr>
        <th valign="top"><label for="email">*E-mail：</label></th>
        <td colspan=2><input name="email" type="text" value="<%=Server.HTMLEncode(email)%>" class="Text" id="email" size="30" maxlength="50"><br/>*供系統認證之用，請務必填寫正確<br/><font color="red">**注意，E-mail 資料一旦變更，您必須重作系統認證程序，以確認此資料的有效性。</font></td>
      </tr>
	  <tr>
        <th><label for="epapercheck">電子報訂閱</label></th>
        <td colspan=2><input id="epapercheck" type="checkbox" name="epapercheck" value="1" <%if checkmail = 1 then%> checked <%END IF%> />訂閱電子報</td>
      </tr>
	  <tr>
        <th><label for="cursorcheck">開啟動態游標</label></th>
        <td colspan=2><input id="cursorcheck" type="checkbox" name="cursorcheck" <%if checkcursor = 1 then%>checked<%END IF%> />是否開啟動態游標</td>
      </tr>
      </table>
	  <!--使用者等級較高者，其擁有的特殊功能權限 Start by Ivy-->	 
	  <input type="hidden" name="level" value="<%=gradeLevel%>">	
	  <%if CInt(gradeLevel) >= 3 then 
	  
	  if right(myWWWsite,1) <> "/" then
	  myWWWsite = myWWWsite & "/"
	  end if
	  %>
	  <BR>	  
	  <table width = "100%" class="sample">
	  <label style="color:blue">此區塊為達人或高手等級擁有的特殊編輯欄位</label>
		<tr>
			<th><label >大頭貼照：</label></th>
			<td style="height:160px;">
				<iframe id="iframeUpload" scrolling="no"  width="500px" frameborder="0" 
					src="<%=myWWWsite%>site/coa/wsxd2/Coamember/UploadAction.aspx?account=<%=account%>">
				</iframe>
					<!--src="http://kmwebsys.coa.gov.tw/site/coa/wsxd2/Coamember/UploadAction.aspx?account=<%=account%>"-->
				
			</td>
		</tr>
		<tr>
			<th><label for="introduction">自我介紹：</label></th>
			<td><textarea id = "introduction" name = "introduction" rows="5" cols="51" onBlur="LimitTextAreaLength(this.form.introduction);"  
       ><%=introduce%></textarea></td>
		</tr>
		<tr>
			<th><label for="contact">個人/公司<BR/>網站或BLOG：</label></th>
			<td><input type="text" size="50" maxlength="255" name="contact" id="contact" value="<%=contact%>" >格式請依循 http://www.xxx.xxx 方式輸入</td>
		</tr>
	  </table>
	  <br>
		
	  <% End If %>
	  <!--使用者等級較高者，其擁有的特殊功能權限 End-->

	  <!--Table CSS-->
<style type="text/css">
table.sample {
	border-width: 5px 5px 5px 5px;
	border-spacing: 2px;
	border-style: double double double double;
	border-color: rgb(99, 177, 255) rgb(99, 177, 255) rgb(99, 177, 255) rgb(99, 177, 255);
	border-collapse: separate;
	background-color: white;
}

</style>
	  <!---->
	    
      <% If SMemberId <> "" Then %>
				<input name="Submit" type="submit" class="Button" value="確定修改">			  
			<% End If %>
			</form>			
				<br>
			</div>
	</div>
</div>
