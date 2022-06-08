<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	ForumID = Request.QueryString("ForumID")
	ForumEs = Request.QueryString("ForumEs")


	SQLCom = "select ForumName from Forum Where ForumID = "& ForumID
 	Set RS = Conn.execute(SQLCom)

	FtypeName = "<b>"& rs("ForumName") &" 文章目錄</b>"

   const cid     = 1	'ArticleID
   const cname   = 2	'Subject
   const cparent = 10	'ArticleParent
   const cChild  = 11	'ArticleChildCount
   const ctime   = 7	'PostDate
   const clevel  = 8	'ArticleLevel
   const poser   = 5	'PostUserName
   const cmail   = 6	'PostEMail

 SQLCom = "select * from ForumArticle Where ForumID = "& ForumID &" Order by ArticleLevel, PostDate Desc"
 Set RS = Conn.execute(SQLCom)

 xSQLCom = "select ArticleID from ForumEssenceArticle Where ForumID = "& ForumID &" Order by ArticleID"
 Set xRS = Conn.execute(xSQLCom)
 If Not xRS.EOF Then xForumEs = xRS.getrows()

%>
<html><head>
<meta http-equiv="pragma: nocache">
<script language=javascript>
 var OpenFolder = "image/NotImage.gif"
 var CloseFolder = "image/NotImage.gif"
 var DocImage = "image/NotImage.gif"
</script>
<script src="ArticleTree.js"></script>
<base target="ForumToc">
<title></title>
</head><body topmargin="3" leftmargin="0">
<center>
<form method="POST" name="Menu">
<table border="0" width="95%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">討論區管理 - <% If ForumEs = "" Then %>刪除文章<% Else %>精華區<% End IF %></td>
    <td align="right"><% If Not rs.EOF And ForumEs = "" Then %><input type="button" value="刪除" class="cbutton" OnClick="ForumDel()"><% End If %></td>
  </tr>
  <tr>
    <td colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
</table>
</form>
</center>
<form name=reg method="POST" action="ForumConfig.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>">
<%
	If not rs.EOF Then
	 NowCount = 0 %>
<script language=javascript>
<%	 Response.write "treeRoot = gFld(""title"", """ & FtypeName & """, """",0,0)" & vbcrlf
	 Do while not rs.EOF
	 NowCount = NowCount + 1
	 LinkName = "<a href=ShowArticle_Toc.asp?ForumID="& ForumID &"&ItemID="& ItemID &"&ArticleID="& rs(cid) &" target=ShowArticle>"& rs(cname) &"</a>"
	  If rs(clevel) = 1 And rs(cChild) > 0 Then
	   Response.Write "N"& rs(cid) &"= insFld(treeRoot, gFld(""ShowArticle"", """& LinkName &" ("& MailSend(rs(cmail),rs(poser)) &" "& rs(ctime) &")"&""", """","& rs(cid) &","& rs(cChild) &","& rs(cparent) &"))" & vbcrlf
	  ElseIF rs(clevel) = 1 And rs(cChild) = 0 Then
	   Response.Write "insDoc(treeRoot, gLnk(""ShowArticle"", """& LinkName &" ("& MailSend(rs(cmail),rs(poser)) &" "& rs(ctime) &")"&""", """","& rs(cid) &","& rs(cChild) &","& rs(cparent) &"))" & vbcrlf
	  ElseIf rs(clevel) > 1 And rs(cChild) > 0 Then
	   Response.Write "N"& rs(cid) &"= insFld(N"& rs(cparent) &", gFld(""ShowArticle"", """& LinkName &" ("& MailSend(rs(cmail),rs(poser)) &" "& rs(ctime) &" Reply)"&""", """","& rs(cid) &","& rs(cChild) &","& rs(cparent)&"))" & vbcrlf
	  Else
	   Response.Write "insDoc(N"& rs(cparent) &", gLnk(""ShowArticle"", """& LinkName &" ("& MailSend(rs(cmail),rs(poser)) &" "& rs(ctime) &" Reply)"&""", """","& rs(cid) &","& rs(cChild) &","& rs(cparent) &"))" & vbcrlf
	  End IF
	 RS.movenext
     Loop %>
     initializeDocument();
</script>
<p align="right"><% If ForumEs = "Y" Then %><input type="button" value="確定" class="cbutton" OnClick="ForumEs()"><% End If %></p>
<%	Else %>
<p></p>
<p>　</p>
<p align="center"><b><font color="#0000FF">此討論板目前無文章！</font></b></p>
<%  End If %>
</form>
</body></html>
<script language="vbscript">
function chsel(xa, b)
<% If NowCount > 1 And ForumEs = "" Then %>
	a = xa -1
	xmyid = Reg.ArticleID(a).value
	xxselectChild xmyid, Reg.ArticleID(a).checked
<% End IF %>
end function

sub xxselectChild(myid, tORf)
	for xxi=0 to Reg.ArticleID.length -1
		if Reg.ArticleID(xxi).parent = myid then
			Reg.ArticleID(xxi).checked = tORf
			xxSelectChild Reg.ArticleID(xxi).value, tORf
		end if
	next
end sub

<% If ForumEs = "Y" Then %>
 Sub ForumEs()
  reg.action="ForumConfig.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>&ConfigCk=ForumEs&ForumEs=<%=ForumEs%>"
  reg.submit
 End Sub

 <% If IsArray(xForumEs) = True Then %>
 Sub window_OnLoad()
  <% NowArray = "Array("
      For xno = 0 to ubound(xForumEs,2)
       NowArray = NowArray & xForumEs(0,xno) & ","
      Next
     NowArray = left(NowArray,instrrev(NowArray,",")-1) & ")"
     response.write "NowArray = " & NowArray %>
    for xxi=0 to Reg.ArticleID.length -1
     for xno = 0 to ubound(NowArray)
      If cint(Reg.ArticleID(xxi).value) = cint(NowArray(xno)) Then
       Reg.ArticleID(xxi).checked = True
      End If
     next
    Next
 End Sub
 <% End IF %>

<% Else %>
 Sub ForumDel()
  reg.action="ForumConfig.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>&ConfigCk=DelA"
  reg.submit
 End Sub
<% End IF %>
</script>
<%
	Function MailSend(xMail,xName)
	 If xMail <> "" Then
	  MailSend = "<a href=mailto:"& xMail &">"& xName &"</a>"
	 Else
	  MailSend = xName
	 End If
	End Function

%>