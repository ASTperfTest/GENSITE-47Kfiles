<%
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

'response.write "**" & session("pwd") & "==<HR>"
'response.end

If Not Session("pwd") Then%>
  <script language=VBScript>
      	Alert ("系統停滯時間逾時！請重新登入！")
    	top.location.href="../default.asp"
  </script>
<%
End If

progPath = request.serverVariables("PATH_INFO")
posPath = InStrRev(progPath, "/")
if posPath > 0 then		progPath = mid(progPath, posPath+1)
posPath = Instr(progPath,".")
if posPath > 0 then		progpath = left(progPath, posPath-1)
	xPos=Instr(session("uGrpID"),",")
	if xPos>0 then
		IDStr=replace(session("uGrpID"),", ","','")
		IDStr="'"&IDStr&"'"
	else
		IDStr="'" & session("ugrpID") & "'"
	end if
   Sql = "SELECT rights FROM ugrpAP INNER JOIN embsysUser ON ugrpAP.ugrpID IN (" & IDStr & ")" & _
         "WHERE embsysUser.userId=N'" & session("UserID") & "' and ugrpAP.APcode=N'" & HTProgCode & "'"
   set RS=conn.execute (Sql)
if rs.eof then %>
  <script language=VBScript>
    Alert ("無使用權限！")
    top.location.href="../default.asp"
  </script>
<%
else
   HTProgRight=0
   Sql = "SELECT APnameC FROM AP Where APCode = '"& HTProgCode &"'"
   set RScom=conn.execute (Sql)
   If Not rscom.EOF Then
   	Title = rscom("APnameC") & "&nbsp;管理"
   	Title2 = rscom("APnameC")
   end if
   while not RS.EOF
   	HTProgRight=(HTProgRight OR rs("rights"))
   	RS.movenext
   wend
'   HTProgRight=rs("rights")
end if

if ((HTProgRight AND 1) = 0) AND (lCase(Right(progPath,5))="query") then
%>
<HTML>
<BODY>
<SCRIPT LANGUAGE=vbs>
msgBox "‧該系統未授權"
window.history.back
</SCRIPT>
</BODY>
</HTML>

<%
   Response.End
End If
%>
