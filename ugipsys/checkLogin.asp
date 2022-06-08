<%@ CodePage = 65001 %>
<% 
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1 
%>

<!-- #INCLUDE FILE="inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->

<%

'// purpose: decode a string
'// ex: ret = unicodeDecode(inputString, publicKey)
function unicodeDecode(byval inputString, byval publicKey)
    dim iPos
    dim returnValue
    dim currentChar

    returnValue = ""
    currentChar = ""

    if inputString <> "" then
        for iPos = 1 to len(inputString)
            currentChar = mid(inputString,iPos,1)
            'response.write "<br>" & iPos & ":" & currentChar & " --> " & uncodDecodeChar(currentChar, publicKey)
            returnValue = returnValue & uncodDecodeChar(currentChar, publicKey)
        next
    end if

    unicodeDecode = returnValue
end function

'// purpose: decode a char
'// ex: ret = unicodeDecodeChar(inputChar, publicKey)
function uncodDecodeChar(byval inputChar, byval publicKey)
    dim iPos
    dim returnValue
    dim currentByte

    returnValue = ""
    currentByte = ""

    if inputChar <> "" then
        for iPos = 1 to lenb(inputChar)
            currentByte = midb(inputChar,iPos,1)
            if iPos=1 then
                currentByte = chrb(ascb(currentByte)-publicKey)
            end if
            returnValue = returnValue & currentByte
        next
    end if
    
    uncodDecodeChar = returnValue
end function

'--start-- 2011/8/25modify by vivian:取得Web.config AppSettingKey
Function GetAppSetting(strAppSettingKey)
  Dim xmlWebConfig
  Dim nodeAppSettings
  Dim nodeChildNode
  Dim strAppSettingValue

  Set xmlWebConfig = Server.CreateObject("Msxml2.DOMDocument.6.0")
  xmlWebConfig.async = False
  xmlWebConfig.Load(Server.MapPath("/Web.config"))

  If xmlWebConfig.parseError.errorCode = 0 Then
    Set nodeAppSettings = xmlWebConfig.selectSingleNode("//configuration/appSettings")
    For Each nodeChildNode In nodeAppSettings.childNodes 
      If nodeChildNode.Text = "" Then
          If nodeChildNode.getAttribute("key") = strAppSettingKey Then
            strAppSettingValue = nodeChildNode.getAttribute("value")
            Exit For
          End If
      End If  
    Next
    Set nodeAppSettings = Nothing
  End If
  Set xmlWebConfig = Nothing

  GetAppSetting = strAppSettingValue
End Function
'--end-- 2011/8/25modify by vivian

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")
	
	if request.form("hidguid") <> "" then
	    sqlCom = "select UserID, Password " _
	        & "from InfoUser " _
	        & "inner join SSO on InfoUser.UserID = SSO.USER_ID " _
	        & "where SSO.GUID = '" & request.form("hidguid") & "'" _
	        & "and SSO.CREATION_DATETIME between dateadd(hh,-2,getdate()) and dateadd(hh,1,getdate()) " 
	    set RS = conn.execute(sqlCom)
	    if not RS.eof then
	        userName = RS("UserID")
	        passWord = RS("Password")
	    else
%>
<HTML>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<BODY>
<script language=vbs>
	MsgBox "帳號/密碼 不合！請重新簽入！"
	top.location.href="login.asp"
</script>
<%	    
	    end if
	else	
	    userName = request.form("tfx_userName")
	    passWord = request.form("tfx_passWord")
    	
	    publicKey = 18
	    userName = unicodeDecode(userName,publicKey)
	    passWord = unicodeDecode(passWord,publicKey)		    
	end if

    '--start-- 2011/8/25modify by vivian:SSO網絡系統帳號登入主題館
    sqlCom = "select * " _
        & "from InfoUser " _
        & " WHERE userId=" & pkStrWithSriptHTML(userName,"")  
    set RS = conn.execute(sqlCom)
    if RS.eof then    
        Dim subject_userName
        url = GetAppSetting("KMintra") + "/coa/login.aspx?sysuserid="+userName+"&syspassword="+passWord
        set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
        xmlhttp.open "GET", url, false 
        xmlhttp.send "" 
        if InStr(xmlhttp.responseText,":success") <> 0 then
            subject_userName = Left(xmlhttp.responseText, InStr(xmlhttp.responseText,":success")-1)
            set xmlhttp = nothing 
            
            if subject_userName <> "fail" then
                sqlCom = "select UserID, Password " _
                & "from InfoUser " _
                & "where UserID = '" & subject_userName & "'" 
                set RS = conn.execute(sqlCom)
                if not RS.eof then
                    userName = RS("UserID")
                    passWord = RS("Password")
                end if
            end if
        end if
	end if
	'--end-- 2011/8/25modify by vivian
	
	if session("DGBAS_uid") <> "" then
		sqlCom = "SELECT i.*, d.tDataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStrWithSriptHTML(session("DGBAS_uid"),"")
	else
		sqlCom = "SELECT i.*, d.tdataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStrWithSriptHTML(userName,"") _
			& " AND password=" & pkStrWithSriptHTML(passWord,"")
	end if	

	set RS = conn.execute(sqlCom)
	if RS.eof then
%>
<HTML>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<BODY>
<script language=vbs>
    MsgBox "帳號/密碼 不合！請重新簽入！"
    window.history.back
</script>
<%
	else
	
  	if checkGIPconfig("UserIPcheck") then
  	
			if isNumeric(RS("NetIP1")) then
		  	raIP = split(request.serverVariables("REMOTE_ADDR"),".")
				ipWrong = False
				for xi = 1 to uBound(raIP) + 1
			  	if isNumeric( RS("NetMask" & xi) ) then
			    	NetMask = cint( RS("NetMask" & xi) )
			    else
			      NetMask = 255
			    end if
					'response.write raIP(xi-1) & ", RIP=" & RS("NetIP"&xi) & ", mask=" & NetMask & "<BR>"
					'response.write (raIP(xi-1) AND NetMask) & "," & (cint(RS("NetIP"&xi)) AND NetMask) & "<BR>"
					if (raIP(xi-1) AND NetMask) <> (cint(RS("NetIP"&xi)) AND NetMask) then	ipWrong = True
				next

				if ipWrong then
%>
					<HTML>
						<BODY>
							<script language=vbs>
								msgBox "IP 位址不合！"
								window.history.back
							</script>
<%
							response.end
				end if
			end if
		end if
	
		if checkGIPconfig("UserLogFile") then
	
			sql = " insert into UserLoginInfo(userid,logintime,loginip,act) " & _
						" values(" & pkStrWithSriptHTML(request.form("tfx_userName"),"") &  ", " & _
						" getdate(), '" & Request.ServerVariables("REMOTE_ADDR") & "', '登入') "
			sql = " set nocount on;" & sql & "; select @@IDENTITY as NewID"
			set RSx = conn.Execute(sql)
			session("loginLogSID") = RSx(0)

		end if

		session("lastVisit") 	= RS("LastVisit")
		session("lastIP") 		= RS("lastIP")
		session("VisitCount") = RS("VisitCount")
		LastVisit = Date()
		VisitCount = RS("VisitCount") + 1
		SQL = " Update InfoUser Set lastVisit = " & pkStr(LastVisit, "") & ", visitCount = " & VisitCount & ", " & _
					" lastIp = '" & request.serverVariables("REMOTE_ADDR") & "' Where userId = " & pkStrWithSriptHTML(RS("UserID"), "") & " "
		
		conn.execute(SQL)
		
		session("pwd") 				= True
		session("userID") 		= stripHTML( RS("userID") )
		session("userName") 	= stripHTML( RS("userName") )
		session("uGrpID") 		= RS("uGrpID")
		session("deptID") 		= RS("deptID")
		session("tDataCat") 	= RS("tDataCat")
		session("uploadPath") = xPath
		If IsNull(RS("Email")) Then
			session("Email") 		= ""
		Else
			session("Email") 		= stripHTML( RS("Email") )
		End If
		
		
		xPath = "/public/data/" & trim( RS("uploadPath") ) 
		if right(xPath,1) <> "/" then	xPath = xPath & "/"
		
		if checkGIPconfig("CheckLoginPassword") then
			session("password") = RS("password")
		end if
		
		if checkGIPconfig("AccountDeny") then
    	if instr(session("uGrpID"), "999") > 0 then
      	session("pwd") = false
        response.write "<script language=javascript>alert('您沒有權限使用此系統！');history.go(-1);</script>"
        response.end
      end if
    end if
				
		if checkGIPconfig("EATWebFormAP") then
			session("EATWebFormAP") = trim(RS("EATWebFormAP") & " ")
		end if
%>
		<script language=vbs>
			<% if Instr(session("uGrpID")&",", "HTSD,") > 0 then %>
					'top.location.href="POPDown.htm"
					top.location.href="default.asp"
			<% else %>
					top.location.href="default.asp"
			<% end if %>
		</script>
	</BODY>
</HTML>
<% end if %>
