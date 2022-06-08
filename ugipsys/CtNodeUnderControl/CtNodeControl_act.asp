<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "NUC" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<%
  '取得輸入
  userId = request("userId")
  ctNodeId = request("ctNodeId")

'response.write userId & "<br>"
'response.write ctNodeId & "<br>"

  '先全部刪除
  SqlCom = "delete from ctNodeUser where userId=N' " & userId & " ' "
  conn.execute(SqlCom)
  
  '再塞入
  MyArray = Split(ctNodeId,",")
  for i=0 to UBound(MyArray)
'response.write MyArray(i) & "<br>"
    SqlCom = "insert into ctNodeUser (userId,ctNodeId,rights) values (N'" & userId & "'," & MyArray(i) & ",1)"
'response.write SqlCom &"<br>"
'response.end
    conn.execute(SqlCom)
  next
  
  '跳回
  Response.Write "<html><body bgcolor='#ffffff'>"
  Response.Write "<script language='javascript'>"
  Response.Write "alert('設定完成!');"
  'Response.Write "location.replace('userList.asp')"
  Response.Write "</script>"
  Response.Write "</body></html>"  
%>
