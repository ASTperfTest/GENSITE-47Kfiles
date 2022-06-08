
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


 'CursorType
      'ForwardOnly--> 0
      'OpenKeyset->1
      'Dynamics-> 2
      'Statics->3
 'LockType
      'ReadOnly --->1
      'Pessim ------>2
      'Optim----->3
      'BatchOptim--->4
 'CursorLocation
      'adUseClient----->1
      'adUseServer----->2   default
 'Rs.open sql,conn,1,1
 'Rs.CurSorLocation=3 --> ClientBatch
 'Rs.CurSorLocation=2 --> Client
 'Rs.CurSorLocation=1 --> Server
%>

<%
If not Session("pwd") Then%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
  <script language=VBScript>
'	document.domain = "corp.com.tw"
      	Alert ("系統停滯時間逾時！請重新登入！")
'    	top.location.href="/default.asp"
    	document.location.href="/reLogin.asp"
  </script>
</body>
</html>
<%
End If 
%>
<%
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
   Sql = "SELECT rights FROM UgrpAp INNER JOIN InfoUser ON UgrpAp.ugrpId IN (" & IDStr & ")" & _
         "WHERE InfoUser.userId=N'" & session("UserID") & "' and UgrpAp.apcode=N'" & HTProgCode & "' Order By rights DESC"       

		'response.write sql 
		'response.end
   set RS=conn.execute (Sql)
if Rs.Eof then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
  <script language=VBScript>
'    document.domain = "corp.com.tw"
    Alert ("無使用權限！")
    top.location.href="../default.asp"
  </script>
</body>
</html>
<%
else
   HTProgRight=0
   Sql = "SELECT apnameC, apcatCname FROM Ap JOIN Apcat ON Ap.apcat=Apcat.apcatId Where apcode = '"& HTProgCode &"'"
   set RScom=conn.execute (Sql)
   If Not rscom.EOF Then
   	Title = rscom("apcatCname") & "／" & rscom("apnameC") 
   	Title2 = rscom("apnameC")
   end if
   while not RS.EOF
   	HTProgRight=(HTProgRight OR rs("rights"))
   	RS.movenext
   wend
'   HTProgRight=rs("rights")
end if

if ((HTProgRight AND 1) = 0) AND (lCase(Right(progPath,5))="query") then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
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
