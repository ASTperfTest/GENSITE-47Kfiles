<%@ CodePage = 65001 %>
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
progPath="D:\hyweb\GENSITE\ugipsys\member\newMemberEdiy_Act.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array()
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array("realnametext","nicknametext")
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array("realnametext","nicknametext")
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################

'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=xUpForm(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "xUpForm(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=xUpForm(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "xUpForm(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=xUpForm(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "xUpForm(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then 
		response.redirect(onErrorPath)
		response.end
	end if
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


Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% 
 
'dim conn 
'Set conn = Server.CreateObject("ADODB.Connection")
'StrCn="Provider=SQLOLEDB;Data Source=10.10.5.127;User ID=hygip;Password=hyweb;Initial Catalog=mGIPcoanew"

'----------HyWeb GIP DB CONNECTION PATCH----------
''conn.Open StrCn
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
'conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString=session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open

'園藝專家
' dim isGardenPro
' isGardenPro = Request("gardenPro")
'response.write Request.QueryString("gardenPro")
'response.write  Request("gardenPro") & "123"
'response.end
' if isGardenPro == "1" then
' expertSql = "INSERT INTO GARDENING_EXPERT "
' expertSql = expertSql &	"(ACCOUNT,INTRODUCTION,SORT_ORDER) VALUES "	
' expertSql = expertSql & "('"& id &"','test',2)"
' conn.execute(expertSql)
' end if

	HTUploadPath = session("Public")
	apath = server.mappath(HTUploadPath) & "\"
	
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage = 65001
		xup.Start apath
		
function xUpForm(xvar)
	xUpForm = trim(xup.form(xvar))
end function

dim id
id = request.querystring("account")

'sam 檢查字串與用的
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage


'園藝專家
dim isGardenPro

isGardenPro = replace(xUpForm("gardenPro"),"'","''")
intro = replace(xUpForm("textfield"),"'","''")
order = replace(xUpForm("order"),"'","''")
'response.write order&"<br>"
'response.write xUpForm("order")
'response.end

if order = "" then order=DBNull end if 

'response.write aaa
'response.end
'有勾園藝專家的時候
if isGardenPro <> "" then
	
	'有資料就update
	'沒有就新增
	expertSql = "select * from GARDENING_EXPERT where account=" & "'" & id&"'"
	set garden_rs = conn.execute(expertSql)

	if not garden_rs.eof then
		'response.write "進入修改"
		'response.end		
		expertSql = "update GARDENING_EXPERT SET "
		expertSql = expertSql  & " INTRODUCTION  =" & "'" & intro &"',"
		expertSql = expertSql  & " SORT_ORDER =" & "'" & order &"'"
		expertSql = expertSql  & " WHERE  ACCOUNT= "& "'"  & id & "'"
		conn.execute(expertSql)
	else
		'response.write "進入新增"
		'response.end
		expertSql = "INSERT INTO GARDENING_EXPERT "
		expertSql = expertSql &	"(ACCOUNT,INTRODUCTION,SORT_ORDER) VALUES "	
		expertSql = expertSql & "('"& id &"','"& intro &"','"& order &"')"
		conn.execute(expertSql)
	end if
else
	'response.write "進入刪除"
	'response.end
	expertSql = "DELETE FROM GARDENING_EXPERT " 
	expertSql = expertSql  & " WHERE  ACCOUNT= "& "'"  & id & "'"
	conn.execute(expertSql)
end if


accounttext = replace(xUpForm("accounttext"),"'","''")
accounttext = trim(accounttext)


Set rs2 = conn.Execute("SELECT COUNT(email) AS count FROM Member WHERE (account <> '" & accounttext & "') AND (email = '" &  trim(replace(xUpForm("emailtext"),"'","''")) & "')")
if not rs2.eof then
	emailCount = rs2("count")
end if	
if emailCount > 0 then
    response.write "<script>alert('E-mail 已被登記使用(" & emailCount  & "次)，請重新輸入');history.back();</script>"
    response.End
end if



photo = replace(xUpForm("imgfile"),"'","''")
accountsql = "SELECT COUNT(*) AS ACCOUNT from Member where account = '" & accounttext & "'"
Set RSreg = conn.execute(accountsql) 


if RSreg("ACCOUNT") = 1 and id <> accounttext then
	response.write "<script>alert('此帳號已註冊');history.back();</script>"
else		
	realnametext 		= replace(xUpForm("realnametext"),"'","''")
	nicknametext 		= xUpForm("nicknametext")
	passwd 				= replace(xUpForm("passwd"),"'","''")
	password2 			= replace(xUpForm("password2"),"'","''")
	idntext 			= replace(xUpForm("idntext"),"'","''")
	member_orgtext 	    = replace(xUpForm("member_orgtext"),"'","''")
	com_teltext 		= replace(xUpForm("com_tel"),"'","''")
	com_exttext			= replace(xUpForm("com_ext"),"'","''")
	ptitletext			= replace(xUpForm("ptitle"),"'","''")
	birthYeartext		= replace(xUpForm("birthYeartext"),"'","''")
	birthMonthtext	    = replace(xUpForm("birthMonthtext"),"'","''")
	birthdaytext		= replace(xUpForm("birthdaytext"),"'","''")
	sextext				= replace(xUpForm("sex"),"'","''")
	homeaddrtext		= replace(xUpForm("homeaddrtext"),"'","''")
	ziptext				= replace(xUpForm("ziptext"),"'","''")
	phonetext			= replace(xUpForm("phonetext"),"'","''")
	home_exttext		= replace(xUpForm("home_exttext"),"'","''")
	mobiletext			= replace(xUpForm("mobiletext"),"'","''")
	faxtext				= replace(xUpForm("faxtext"),"'","''")
	emailtext			= replace(xUpForm("emailtext"),"'","''")
	emailverified		= replace(xUpForm("emailverified"),"'","''")
	'status=replace(xUpForm("status"),"'","''")
	upLoadRight			= replace(xUpForm("upLoadRight"),"'","''")
	uploadPicCount	    = replace(xUpForm("uploadPicCount"),"'","''")
	scholarValidate=replace(xUpForm("scholarValidate"),"'","''")
	id_type2=replace(xUpForm("id_type2"),"'","''")
	knowledgePro		= replace(xUpForm("knowledgePro"),"'","''")
	KMcat				= replace(xUpForm("KMcat"),"'","''")
	ShowCursorIcon	    = replace(xUpForm("ShowCursorIcon"),"'","''")'0812 Grace
	remark			    = replace(xUpForm("remark"),"'","''")'0812 Grace

	if knowledgePro="1" then
		knowledgePro="1"
	else
		knowledgePro="0"
	end if
	
	birthday = birthYeartext & birthMonthtext & birthdaytext

sql = "UPDATE Member SET  "
sql = sql  & "account =" & "'" & accounttext &"'"
sql = sql  & ",realname =" & "'" & realnametext &"'"
sql = sql  & ",nickname =" & "'" & nicknametext &"'"
sql = sql  & ",passwd = "& "'" & passwd  &"'"
sql = sql  & ",id = "& "'" & idntext  &"'"
sql = sql  & ",member_org = "&"'" & member_orgtext  &"'"
sql = sql  & ",com_tel = "&"'" & com_teltext  &"'"
sql = sql  & ",com_ext = "&"'" & com_exttext  &"'"
sql = sql  & ",ptitle = "& "'" & ptitletext  &"'"
sql = sql  & ",birthday = "& "'" & birthday  &"'"
sql = sql  & ",sex = "& "'" & sextext  &"'"
sql = sql  & ",homeaddr = "& "'" & homeaddrtext  &"'"
sql = sql  & ",zip = "& "'" & ziptext  &"'"
sql = sql  & ",phone = "&"'" & phonetext  &"'"
sql = sql  & ",home_ext = "& "'" & home_exttext  &"'"
sql = sql  & ",mobile = "& "'" & mobiletext  &"'"
sql = sql  & ",fax = "&"'" & faxtext  &"'"
sql = sql  & ",email = "&"'" & emailtext  &"'"
If emailverified = "Y" then
sql = sql  & ",mcode = "&"'" & emailverified  &"'"
Else
sql = sql  & ",mcode = NULL"
End if
'sql = sql  & ",status = "&"'" & status  &"'"
sql = sql  & ",upLoadRight = "&"'" & upLoadRight  &"'"
sql = sql  & ",uploadPicCount = "&"'" & uploadPicCount  &"'"

'新增功能:可直接改為學者會員身分  
scholarcheck = replace(xUpForm("scholarcheck"),"'","''") 
if scholarcheck = "1" then
    sql = sql  & ",scholarValidate = "&"'Y'"
	sql = sql  & ",id_type2 = "&"'1'"
end if
'sql = sql  & ",scholarValidate = "&"'" & scholarValidate  &"'"
sql = sql  & ",id_type3 = "&"'" & knowledgePro  &"'"
sql = sql  & ",KMcat = "&"'" & KMcat  &"'"
sql = sql  & ",remark = "&"'" & remark  &"'"'0812 Grace

'開啟動態游標處理 Grace
cursorcheck = replace(xUpForm("cursorcheck"),"'","''")
if cursorcheck = "1" then
	sql = sql  & ",ShowCursorIcon = "&"'1'"'0812 Grace
else
	sql = sql  & ",ShowCursorIcon = "&"'0'"'0812 Grace
end if
'開啟動態游標處理 end




if  trim(photo)<>"" then
	'nfname = Year(now()) & Month(now()) & day(now()) & Hour(now())& Minute(now()) & Second(now()) 
	nfname = xup.Form("imgfile").FileName
	pos = instr(nfname, ".")
	nfname = Year(now()) & Month(now()) & day(now()) & Hour(now())& Minute(now()) & Second(now()) &"." & mid(nfname, pos + 1)
	route =  "\"& "public" & "\" & nfname
	xup.Form("imgfile").SaveAs apath & nfname, True	
	sql = sql  & ",photo = "&"'" & route  &"'"
end if 

sql = sql  & " WHERE account = " & "'" & id&"'"	


conn.execute(sql)



'訂閱電子報處理
epapercheck = replace(xUpForm("epapercheck"),"'","''")
if epapercheck = "1" then
	checksql = "select * from Epaper where email = '"&emailtext&"'"
	set check_epaper = conn.execute(checksql)
	if check_epaper.eof then
		sql1 = "INSERT INTO Epaper (email,createtime,CtRootID) VALUES ('"&emailtext&"', getdate(),'21')"
		conn.execute(sql1)
	end if
else
	sql2 = "DELETE FROM Epaper WHERE email = '"& emailtext &"'"
	conn.execute(sql2)
end if
'訂閱電子報處理end



'存檔後回原頁 Grace
if request.QueryString("from") = "list" then
    response.redirect "newMemberList.asp?keep=Y&validate="&session("validate")&"&nowPage="&request.QueryString("nowPage")&"&pagesize="&request.QueryString("pagesize")
elseif request.QueryString("from")  = "query" then
    response.redirect "newMemberQuery_Act.asp?nowPage="&request.QueryString("nowPage")&"&pagesize="&request.QueryString("pagesize")
else
    response.redirect "newMemberList.asp?keep=Y&validate="&session("validate")
end if
'存檔後回原頁 end


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
