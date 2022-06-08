<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<%
  '取得輸入
  userId = request("userId")
  CtNodeID = request("CtNodeID")

  '先全部刪除
  SqlCom = "delete from CtUserSet where userId=N'" & userId & "'"
  conn.execute(SqlCom)
  
  '再塞入
  MyArray = Split(CtNodeID,",")
  for i=0 to UBound(MyArray)
'response.write MyArray(i) & "<br>"
    SqlCom = "insert into CtUserSet (userId,ctNodeId,rights) values (N'" & userId & "'," & MyArray(i) & ",1)"
    conn.execute(SqlCom)
  next
  
  '跳回
  Response.Write "<html><body bgcolor='#ffffff'>"
  Response.Write "<script language='javascript'>"
  Response.Write "alert('設定完成!');"
  Response.Write "location.replace('userList.asp')"
  Response.Write "</script>"
  Response.Write "</body></html>"  
%>