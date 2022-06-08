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
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	if session("DGBAS_uid") <> "" then
		sqlCom = "SELECT i.*, d.tDataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStrWithSriptHTML(session("DGBAS_uid"),"")
	else
		sqlCom = "SELECT i.*, d.tdataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStrWithSriptHTML(request.form("tfx_userName"),"") _
			& " AND password=" & pkStrWithSriptHTML(request.form("tfx_passWord"),"")
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

	
'		if checkGIPconfig("MMOFoler") then
''    session("MMOPublic") = "/site/coa/public/"	
'		end if

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
