<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
	ForumID = Request.QueryString("ForumID")

		 SQLCom = "Select * from Forum Where ForumID = "& ForumID
		 Set RS = Conn.execute(SQLCom)
			opinion=rs("ForumRemark")
			cont=message(opinion)
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>討論板清單</title>
<link rel="stylesheet" href="../inc/setstyle.css">
</head><body>
<center>
<form method="POST" name="Menu">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="FormName">討論區管理 - 檢視討論板</td>
    <td align="right"><% IF ForumType <> "D" Then %><input type="button" value="編修" class="cbutton" OnClick="ForumEdit()"><% End IF %><input type="button" value="刪除文章" class="cbutton" OnClick="ArticleEdit()"><input type="button" value="精華區管理" class="cbutton" OnClick="ArticleEsEdit()"></td>
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
<table border="0" width="80%" cellspacing="1" cellpadding="3" class="CTable">
  <tr>
    <td align="right" width="30%" class="TableHeadTD">討論板名稱</td>
    <td class="TableBodyTD"><%=rs("ForumName")%></td>
  </tr>
  <tr>
    <td align="right" width="30%" class="TableHeadTD">討論板說明</td>
    <td class="TableBodyTD"><%=cont%></td>
  </tr>
</table>
</center>
</body></html>
<script language="VBScript">
	Sub ForumEdit()
	 window.location.href = "ForumEdit.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>"
	End Sub

	Sub ArticleEdit()
	 window.location.href = "ArticleView.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>"
	End Sub

	Sub ArticleEsEdit()
	 window.location.href = "ArticleView.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>&ForumEs=Y"
	End Sub
</script>
<%
function message(tempstr)
  outstring = ""
  while len(tempstr) > 0
    pos=instr(tempstr, chr(13)&chr(10))
    if pos = 0 then
      outstring = outstring & tempstr & "<br>"
      tempstr = ""
    else
      outstring = outstring & left(tempstr, pos-1) & "<br>"
      tempstr=mid(tempstr, pos+2)
    end if
  wend
  message = outstring
end function
%>