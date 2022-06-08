<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	CatID = Request.QueryString("CatID")
	NotDataText = "無子目錄"
	NotDataCk = "N"

	IF CatID = 0 Then
	 Sql = "SELECT 0 AS ctNodeId, ctRootName AS catName, editDate" _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='C' AND dataParent=0 " _
	 	& " AND ctRootId = " & ItemID & ") AS xChildCount" _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='U' AND dataParent=0 " _
	 	& " AND ctRootId = " & ItemID & ") AS xDataCount" _
	 	& " FROM CatTreeRoot WHERE ctRootId=" & pkStr(ItemID,"")
	 	dataLevel = 0
'	 response.write sql
'	 response.end
	 set RS = conn.execute(Sql)
	 SQLCom = "Select htx.* " _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='C' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xChildCount" _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='U' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xDataCount" _
	 	& " FROM CatTreeNode AS htx Where ctNodeKind='C' AND ctRootId = N'"& ItemID &"' And dataLevel = 1 Order by catShowOrder"
	 Set RSCom = Conn.execute(SQLCom)
	Else
	 SQL = "Select htx.* " _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='C' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xChildCount" _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='U' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xDataCount" _
	 	& " FROM CatTreeNode AS htx Where ctNodeId = "& CatID
	 Set RS = Conn.execute(SQL)
	 	dataLevel = RS("dataLevel")
	 SQLCom = "Select htx.* " _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='C' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xChildCount" _
	 	& ", (SELECT count(*) FROM CatTreeNode WHERE ctNodeKind='U' AND dataParent=htx.ctNodeId " _
	 	& " AND ctRootId = " & ItemID & ") AS xDataCount" _
	 	& " FROM CatTreeNode AS htx Where ctNodeKind='C' AND dataParent = "& CatID &" Order by catShowOrder"
	 Set RSCom = Conn.execute(SQLCom)
	 If rs.EOF And RsCom.EOF Then
	  NotDataText = "目錄已刪除選擇其他目錄"
	  NotDataCk = "Y"
	 End If
	End If

%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>目錄清單</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">目錄樹管理 - 目錄清單</td>
    <td align="right">
    <% If NotDataCk = "N" Then %>
    <input type="button" value="新增子目錄" class="cbutton" OnClick="AddCat()"> &nbsp; 
    <input type="button" value="新增項目" class="cbutton" OnClick="AddForum()">
    <% End IF %></td>
  </tr>
  <tr>
    <td colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td colspan="2" class="Formtext"></td>
  </tr>
</table>
</form>
<% IF CatID <> 0 Then
    IF Not rs.EOF Then
      PForumName = rs("CatName")
      dataLevel = rs("dataLevel") + 1
      DataTop = 0
      dataParent = rs("ctNodeId")
  	   DelCk = ""
  	   If RS("xDataCount") = 0 And RS("xChildCount") = 0 Then DelCk = "Y" %>
<table border="0" width="90%" cellspacing="1" cellpadding="3" class="ctable">
  <tr class="TableHeadTD">
    <td align="center" width="50%">目錄名稱</td>
    <td align="center" width="15%">項目個數</td>
    <td align="center" width="15%">子目錄個數</td>
    <td align="center" width="20%">編修日期</td>
  </tr>
  <tr class="TableBodyTD">
    <td><a href="CatView.asp?CatID=<%=RS("ctNodeId")%>&ItemID=<%=ItemID%>&DelCk=<%=DelCk%>"><%=rs("CatName")%></a></td>
    <td align="center"><%=rs("xDataCount")%></td>
    <td align="center"><%=rs("xChildCount")%></td>
    <td><%=d7date(rs("editDate"))%></td>
  </tr>
  <% End IF
    Else
  	  PForumName = "根目錄"
  	  dataLevel = 1
      DataTop = 0
      dataParent = 0 %>
<table border="0" width="90%" cellspacing="1" cellpadding="3" class="ctable">
  <tr class="TableHeadTD">
    <td align="center" width="50%">目錄名稱</td>
    <td align="center" width="15%">項目個數</td>
    <td align="center" width="15%">子目錄個數</td>
    <td align="center" width="20%">編修日期</td>
  </tr>
  <tr class="TableBodyTD">
    <td><%=PForumName%></td>
    <td align="center"><%=rs("xDataCount")%></td>
    <td align="center"><%=rs("xChildCount")%></td>
    <td><%=d7date(rs("editDate"))%></td>
  </tr>
<%	End If
	If Not RSCom.EOF Then %>
  <tr class="TableBodyTD">
    <td colspan="4" class="TableHeadTD">以下為 <%=PForumName%> 子目錄</td>
  </tr>
<% 	  Do while not RSCom.EOF
  	   DelCk = ""
  	   If RSCom("xDataCount") = 0 And RSCom("xChildCount") = 0 Then DelCk = "Y" %>
  <tr class="TableBodyTD">
    <td>　<a href="CatView.asp?CatID=<%=RSCom("ctNodeId")%>&ItemID=<%=ItemID%>&DelCk=<%=DelCk%>"><%=RSCom("CatName")%></a></td>
    <td align="center"><%=RSCom("xDataCount")%></td>
    <td align="center"><%=RSCom("xChildCount")%></td>
    <td><%=d7date(RSCom("editDate"))%></td>
  </tr>
   <% RSCom.movenext
      Loop
     Else %>
  <tr class="TableBodyTD">
    <td colspan="4" align="center"><font color=red><%=NotDataText%></font></td>
  </tr>
<%   End IF %>
</table>
</center>

<%if (HTProgRight and 1)=1 then%>
	<P align="center">
			來源：<SELECT name=orgTree size=1>
<%
	sql = "SELECT * FROM CatTreeRoot WHERE ctRootId<>" & pkStr(request.queryString("ItemID"),"")
	set RSx = conn.execute(sql)
	while not RSx.eof
		response.write "<OPTION VALUE=""" & RSx("ctRootId") & """>" & RSx("CtRootName") & "</OPTION>" & vbCRLF
		RSx.moveNext
	wend
%>
			</SELECT>
			<A id="copyTree" href="#" title="複製目錄樹">複製目錄樹</A>
	</P>
<%end if%>

</body>
<script language=VBScript>
	Sub AddForum()
	 window.location.href = "CtUnitNodeAdd.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&ctNodeKind=U&phase=add"
	End Sub

	Sub AddCat()
	 window.location.href = "CtUnitNodeAdd.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>&ctNodeKind=C&phase=add"
	End Sub

	Sub OrderCat()
	 window.location.href = "CatOrder.asp?ItemID=<%=ItemID%>&CatID=<%=CatID%>"
	End Sub

  sub copyTree_onClick	
  	window.navigate "CtNodeCopy.asp?<%=request.queryString%>&dataLevel=<%=dataLevel%>&orgTree=" & document.all("orgTree").value
  end sub
</script>








