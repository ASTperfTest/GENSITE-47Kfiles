<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	ForumID = Request.QueryString("ForumID")

		 SQLCom = "Select * from Forum Where ForumID = "& ForumID
		 Set RS = Conn.execute(SQLCom)
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>編修討論板</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head>
<body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">討論區管理 - 編修討論板</td>
    <td align="right"><input type="button" value="刪除" class="cbutton" OnClick="Del()"><input type="button" value="檢視" class="cbutton" OnClick="VBS: history.back"></td>
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
<form method="POST" action="ForumConfig.asp?ItemID=<%=ItemID%>&ConfigCk=Edit" name="AddForum">
<input type="hidden" name="CatID" value="<%=rs("CatID")%>">
<input type="hidden" name="ForumID" value="<%=rs("ForumID")%>">
<input type="hidden" name="ForumMaster" value="0">
<table border="0" width="80%" cellspacing="1" cellpadding="3" class="CTable">
  <tr>
    <td align="right" width="30%" class="TableHeadTD">討論板名稱</td>
    <td class="TableBodyTD"><input type="text" name="ForumName" size="20" value="<%=rs("ForumName")%>"></td>
  </tr>
  <tr>
    <td align="right" width="30%" class="TableHeadTD">討論板說明</td>
    <td class="TableBodyTD"><textarea rows="6" name="ForumRemark" cols="45"><%=rs("ForumRemark")%></textarea></td>
  </tr>
</table>
  <table border="0" cellspacing="1" cellpadding="3" width="80%">
    <tr>
      <td align="right"><input type="button" value="確定" class="cbutton" OnClick="FormSubmit()"><input type="button" value="重填" OnClick="VBS: AddForum.reset" class="cbutton">
      </td>
    </tr>
  </table>
</form>
</center>
</body></html>
<script language="VBScript">
 Sub FormSubmit()
  If trim(AddForum.ForumName.value) = Empty Then
     MsgBox "請輸入討論板名稱！", 16, "Sorry!"
     AddForum.ForumName.focus
     Exit Sub
  ElseIf blen(AddForum.ForumName.value) > 30 Then
     MsgBox "討論板名稱字數過多！", 16, "Sorry!"
     AddForum.ForumName.focus
     Exit Sub
  End IF
  AddForum.Submit
 End Sub

 function blen(xs)
 xl = len(xs)
 for i=1 to len(xs)
  if asc(mid(xs,i,1))<0 then xl = xl + 1
 next
 blen = xl
end function

 Sub Del()
  chky=msgbox("您要刪除此討論板及所屬文章!"& vbcrlf & vbcrlf &"　確定嗎？"& vbcrlf , 48+1, "請注意！！")
   if chky<>vbok then
    Exit Sub
   End If
	 AddForum.action="ForumConfig.asp?ItemID=<%=ItemID%>&ConfigCk=Del"
	 AddForum.Submit
 End Sub
</script>