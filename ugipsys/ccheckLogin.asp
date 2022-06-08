<%@ CodePage = 65001 %><% Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 %>
<!-- #INCLUDE FILE="inc/dbutil.inc" -->
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

	sqlCom = "SELECT i.*, d.tDataCat " _
		& " FROM InfoUser AS i LEFT JOIN dept AS d ON i.deptID=d.deptID " _
		& " WHERE userID=" & pkStr(request.form("tfx_userName"),"") _
		& " AND password=" & pkStr(request.form("tfx_passWord"),"")
'	response.write sqlCom
'	response.end ' or 'a'='a
	set RS = conn.execute(sqlCom)
	if RS.eof then
%>
<HTML>
<BODY>
<script language=vbs>
	msgBox "帳號/密碼 不合！請重新簽入！"
	window.history.back
</script>
<%
    response.End
	else
		session("lastVisit") = RS("LastVisit")
		session("lastIP") = RS("lastIP")
		session("VisitCount") = RS("VisitCount")
		LastVisit = Date()
		VisitCount = RS("VisitCount") + 1
		SQL = "Update InfoUser Set LastVisit = '"& LastVisit &"', VisitCount = "& VisitCount _
			& ", lastIP='" & request.serverVariables("REMOTE_ADDR") &"' Where UserID = '"& RS("UserID") &"'"
		conn.execute(SQL)
		session("pwd") = True
		session("userID") = RS("userID")
		session("userName") = RS("userName")
		session("uGrpID") = RS("uGrpID")
		session("deptID") = RS("deptID")
		session("tDataCat") = RS("tDataCat")
		xPath = "/public/data/" & trim(RS("uploadPath")) 
		if right(xPath,1)<>"/" then	xPath = xPath & "/"
		session("uploadPath") = xPath
		session("Email") = RS("Email")
		
		
		'判斷是否有記錄密碼更新日期-------------------------------------------------------------------Start	
		xsql="sp_columns @table_name = 'InfoUser' , @column_name ='ModifyPassword'"
		set xRS= conn.execute(xsql)
	
	  session("ModifyPassword") = False
		if not xRS.eof then 	
			if DateAdd("d",90,RS("ModifyPassword")) < Now then
				session("ModifyPassword") = True
			%>
			<script language=vbs>
				msgBox "密碼超過90天，請修改密碼！"
			</script>
		<%
			end if
		end if
		'判斷是否有記錄密碼更新日期---------------------------------------------------------------------End
%>
<script language=vbs>
<% 	if Instr(session("uGrpID")&",", "HTSD,") > 0 then %>
'		top.location.href="POPDown.htm"
		top.location.href="cdefault.asp"
<%	else %>
		top.location.href="cdefault.asp"
<%	end if %>
</script>
</BODY>
</HTML>
<%
	end if
%>
