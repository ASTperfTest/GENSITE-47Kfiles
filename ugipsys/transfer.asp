
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>管理</title>
</head>
<body >


<% Response.Expires = 0 
   HTProgCode = "HTP02"  
   uid = Session("UserID")

   dPath = Replace(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/"),"\","\\")
   wPath = session("coasiteURL") &"/Public/"
  
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
<%
else
    Response.Redirect "./project0516/transfer.aspx?DBUSERID="& uid & "&dGipDsdPath=" & dPath & "&wPublicPath="& wPath
end if
%>
</body>
</html> 
 
