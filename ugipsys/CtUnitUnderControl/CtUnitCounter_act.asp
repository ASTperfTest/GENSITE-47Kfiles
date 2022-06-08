<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "UWC" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<%'="submitset"& request("submitset")%>
<%'="submitreset"& request("submitreset")%>

<%
  '取得輸入
  userId = request("userId")
  'userId = "hyweb"
  ctUnitId = request("ctUnitId")

'response.write userId & "<br>"
'response.write ctUnitId & "<br>"

	If request("submitSet") <> "" Then
		  '先全部刪除
		  SqlCom = "delete from CtUnitCount where userId='" & userId & "'"
		  conn.execute(SqlCom)
		  
		  '再塞入
		  MyArray = Split(ctUnitId,",")
		  for i=0 to UBound(MyArray)
		'response.write MyArray(i) & "<br>"
		    SqlCom = "insert into CtUnitCount (userId,ctUnitId,rights) values (N'" & userId & "'," & MyArray(i) & ",1)"
		'response.write SqlCom &"<br>"
		'response.end
		    conn.execute(SqlCom)
		  next
	End If
	
	If request("submitReset") <> "" Then
		'先全部刪除
		SqlCom = "delete from CounterCpSetUnit where ctUnitId IN (" & ctUnitId  & ")"
		conn.execute(SqlCom)
	End If
  
  '跳回
  Response.Write "<html><body bgcolor='#ffffff'>"
  Response.Write "<script language='javascript'>"
  Response.Write "alert('設定完成!');"
  'Response.Write "location.replace('userList.asp')"
  Response.Write "</script>"
  Response.Write "</body></html>"  
%>
