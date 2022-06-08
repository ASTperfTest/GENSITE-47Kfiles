<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% ForumID = Request.QueryString("ForumID")
   ArticleID = Request.QueryString("ArticleID") %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head><body topmargin="0">
<% If ArticleID <> "" Then
	 SQL = "SELECT * FROM ForumArticle Where ArticleID = "& ArticleID
 	 Set RS = Conn.execute(SQL)
		opinion=rs("Content")
		cont=message(opinion) %>
<center>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="33"><b><%=rs("Subject")%></b></td>
    <td align="right">
    </td>
  </tr>
</table>
<table border="0" width="90%" cellspacing="1" cellpadding="3" class="ctable">
  <tr class="FormTitleText">
    <td>發表人：<%=MailSend(rs("PostEMail"),rs("PostUserName"))%></td>
    <td align="right"><%=rs("PostDate")%></td>
  </tr>
  <tr class="FormTitleText">
    <td colspan="2" class="TableBodyTD"><%=cont%><br><hr width="85%" noshade size="1" color="#000080">
    <% SQLFile = "Select * FROM FileUp Where ItemID =N'"& ItemID &"' And ParentID = "& ArticleID
       Set RSFile = Conn.execute(SQLFile)
        If Not RSFile.EOF Then %>
      <table border="0" width="100%" cellspacing="0" cellpadding="3" class="TableBodyTD">
        <tr>
          <td width="10%" align="right" valign="top">附件：</td>
          <td width="90%">
          <% Do while not RSFile.EOF
              response.write FileCk(RSFile("NFileName"),RSFile("OFileName"))
             RSFile.movenext
		     Loop %>
          </td>
        </tr>
      </table>
    <%  End IF %>
    </td>
  </tr>
</table>
&nbsp;
</center>
<% End If %>
</body></html>
<%
	Function MailSend(xMail,xName)
	 If xMail <> "" Then
	  MailSend = "<a href=mailto:"& xMail &">"& xName &"</a>"
	 Else
	  MailSend = xName
	 End If
	End Function

	Function FileCk(nfname,ofname)
	  if instr(nfname, ".")>0 then
	    fnext=mid(nfname, instr(nfname, "."))
	    If lcase(fnext) = ".gif" or lcase(fnext) = ".jpg" Then
	     FileCk = "<br><img border=0 src="& session("Public") & nfname &"><br>"
	    Else
	     FileCk = "<a href="& session("Public") & nfname &" target=_blank>"& ofname &"</a><br>"
	    End IF
	  end if
	End Function

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