<%@ CodePage = 65001 %>
<% 
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1 
%>
<!-- #INCLUDE FILE="inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	'931116 Start Add By Jason
	'For DGBAS SSO
	'當有使用者帳號，直接檢查帳號，找出權限即可
	If session("DGBAS_uid")<>"" Then
		sqlCom = "SELECT i.*, d.tDataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStr(session("DGBAS_uid"),"")
	Else
		sqlCom = "SELECT i.*, d.tdataCat " _
			& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
			& " WHERE userId=" & pkStr(request.form("tfx_userName"),"") _
			& " AND password=" & pkStr(request.form("tfx_passWord"),"")
	End If	
	'931116 End Add By Jason
	
	'response.write sqlCom
	'response.end ' or 'a'='a

	Set RS = conn.execute(sqlCom)
	If RS.eof Then
%>
		<HTML>
			<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
			<BODY>
				<script language=vbs>
					MsgBox "帳號/密碼 不合！請重新簽入！"
					window.history.back
				</script>
<%
	Else

  		If checkGIPconfig("UserIPcheck") Then
			If isNumeric(RS("NetIP1")) Then
		  		'response.write Request.ServerVariables("REMOTE_ADDR") & "<br>"
				raIP = split(request.serverVariables("REMOTE_ADDR"),".")
				ipWrong = False
				For xi = 1 To uBound(raIP)+1
			        If isNumeric(RS("NetMask"&xi)) Then
			           NetMask=cint(RS("NetMask"&xi))
			        Else
			           NetMask=255
			        End If
					'response.write raIP(xi-1) & ", RIP=" & RS("NetIP"&xi) & ", mask=" & NetMask & "<BR>"
					'response.write (raIP(xi-1) AND NetMask) & "," & (cint(RS("NetIP"&xi)) AND NetMask) & "<BR>"
					If (raIP(xi-1) AND NetMask) <> (cint(RS("NetIP"&xi)) AND NetMask) then	ipWrong = True
				Next

				If ipWrong Then
%>
					<HTML>
						<BODY>
							<script language=vbs>
								msgBox "IP 位址不合！"
								window.history.back
							</script>
<%
					Response.End
				End If
			End If
		End If
	
		If checkGIPconfig("UserLogFile") Then
			sql="insert into UserLoginInfo(userid,logintime,loginip,act) values(" & pkStr(request.form("tfx_userName"),"") & ",getdate(),'" & Request.ServerVariables("REMOTE_ADDR") & "','登入')"
			sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
			set RSx = conn.Execute(sql)
			session("loginLogSID") = RSx(0)
		End If

		session("lastVisit") = RS("LastVisit")
		session("lastIP") = RS("lastIP")
		session("VisitCount") = RS("VisitCount")
		LastVisit = Date()
		VisitCount = RS("VisitCount") + 1
		SQL = "Update InfoUser Set lastVisit = '"& LastVisit &"', visitCount = "& VisitCount _
			& ", lastIp='" & request.serverVariables("REMOTE_ADDR") &"' Where userId = '"& RS("UserID") &"'"
		conn.execute(SQL)
		session("pwd") = True
		session("userID") = RS("userID")
	
		If checkGIPconfig("CheckLoginPassword") Then
			session("password") = RS("password")
		End If
		
		session("userName") = RS("userName")
		session("uGrpID") = RS("uGrpID")
        
        If checkGIPconfig("AccountDeny") Then
        	If instr(session("uGrpID"), "999") > 0 Then
            	session("pwd") = false
                Response.write "<script language=javascript>alert('您沒有權限使用此系統！');history.go(-1);</script>"
                Response.End
            End If
        End If
		
		session("deptID") = RS("deptID")
		session("tDataCat") = RS("tDataCat")
		xPath = "/public/data/" & trim(RS("uploadPath")) 
		If right(xPath,1) <> "/" Then	xPath = xPath & "/"
		session("uploadPath") = xPath
		session("Email") = RS("Email")
		If checkGIPconfig("EATWebFormAP") Then
			session("EATWebFormAP") = trim(RS("EATWebFormAP") & " ")
		End If
	
		session.Timeout = 300
%>
		<script language=vbs>
		<% 	If Instr(session("uGrpID")&",", "HTSD,") > 0 Then %>
			'top.location.href="POPDown.htm"
			top.location.href="default_rdecthinktank.asp"
		<%
			'response.redirect "login.aspx?usrid=" & session("userID")
			Else 
		%>
			top.location.href="default_rdecthinktank.asp"
		<%	End If %>
		</script>
		</BODY>
	</HTML>
<%
	End If
%>
