<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
	if checkGIPconfig("CheckLoginPassword") and session("userID") = session("password") then
		response.write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
        	response.write "<script language=javascript>alert('請變更密碼、更新負責人員姓名！');location.href='/user/userUpdate.asp';</script>"
        	response.end
	end if
 
 
 
 
 
	If ForumType = "D" Then
	 CatalogueFrame = "0,*"
   Else
	 CatalogueFrame = "220,*"
   End IF

   If ArticleType = "A" Then
	 ForumTocFrame = "35,40%,*"
   Else
	 ForumTocFrame = "0,*,0"
   End IF
   
'   ItemID = request("CtRootID")
   ItemID = session("exRootID")
   tocURL = "Folder_toc.asp"
' 	tocURL =  "xuToc.asp"
  if Instr(session("uGrpID")&",", "HTSD,") > 0 then 
  	tocURL =  "xmlToc.asp"
'  	tocURL =  "xToc.asp"
'  	tocURL =  "xmlToc.asp"
  end if
'response.write "xx=" & tocURL
'response.end
'response.write "<!--" & tocURL & "-->"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<%
sql=""
sql=sql & vbcrlf & "declare @user nvarchar(50)"
sql=sql & vbcrlf & "set @user = '" & Session("UserID") & "'"
sql=sql & vbcrlf & ""
sql=sql & vbcrlf & "select "
sql=sql & vbcrlf & "   CuDTGeneric.iCUItem"
sql=sql & vbcrlf & "FROM CuDTGeneric "
sql=sql & vbcrlf & "INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID "
sql=sql & vbcrlf & "INNER JOIN CatTreeRoot "
sql=sql & vbcrlf & "INNER JOIN NodeInfo ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql=sql & vbcrlf & "INNER JOIN KnowledgeToSubject ON CatTreeRoot.CtRootID = KnowledgeToSubject.subjectId "
sql=sql & vbcrlf & "INNER JOIN CatTreeNode ON KnowledgeToSubject.ctNodeId = CatTreeNode.CtNodeID ON CtUnit.CtUnitID = CatTreeNode.CtUnitID "
sql=sql & vbcrlf & "WHERE (NodeInfo.owner = @user)"
sql=sql & vbcrlf & "AND (KnowledgeToSubject.status = 'Y') "
sql=sql & vbcrlf & "AND (CuDTGeneric.fCTUPublic = 'P')"
sql=sql & vbcrlf & "--不為系統管理員"
sql=sql & vbcrlf & "and not exists(select UserId from infouser where UserId= @user and charindex('SysAdm',ugrpID) > 0)"

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = session("ODBCDSN")
Conn.CursorLocation = 3
Conn.open
set rs = conn.execute (sql)

if rs.recordcount>0 then
%>
    <script type="text/javascript">
        alert('您有新聞尚未審核是否發佈到主題館，請您先審核新聞')
        window.location.href = "/publish_new/NewsApproveList.asp?T=<%=now() %>"
    </script>
<%else%>
<frameset cols="<%=CatalogueFrame%>" framespacing="1" border="1" id="CatalogueMenu">
  <frame name="Catalogue" target="ForumToc" scrolling="auto" noresize src="<%=tocURL%>?ItemID=<%=ItemID%>&branch=<%=request("branch") %>">
    <frame name="ForumToc" scrolling="auto" target="ShowArticle">
  <noframes>
  <body>
  <p>[,Ozs.</p>
  </body>
  </noframes>
</frameset>
<%end if %>
</html>
