<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "HT001" 
' ============= Modified by Chris, 2006/09/12, to handle 上稿權限能以使用者設定 聯集 群組設定 ========================'
'		Document: 950912_智庫GIP擴充.doc
'  modified list:
'	新增此程式
%>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<%
  '取得輸入
  ugrpID = request("ugrpID")
  aType = request("aType")
  CtNodeID = request("CtNodeID")

  '先全部刪除
  SqlCom = "delete from CtUGrpSet where ugrpId=" & pkStr(ugrpId,"") _
	& " AND aType=" & pkStr(aType,"")
  conn.execute(SqlCom)
  
  '再塞入
  MyArray = Split(CtNodeID,",")
  for i=0 to UBound(MyArray)
'response.write MyArray(i) & "<br>"
    SqlCom = "insert into CtUgrpSet (aType,ugrpId,ctNodeId,rights) values (" _
		& pkstr(aType,",") _
		& pkStr(ugrpID,",") _
		& MyArray(i) & ",1)"
    conn.execute(SqlCom)
  next
  
  '跳回
  Response.Write "<html><body bgcolor='#ffffff'>"
  Response.Write "<script language='javascript'>"
  Response.Write "alert('設定完成!');"
  Response.Write "location.replace('listGroup.asp')"
  Response.Write "</script>"
  Response.Write "</body></html>"  
%>