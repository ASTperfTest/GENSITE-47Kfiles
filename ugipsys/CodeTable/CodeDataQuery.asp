<%@ CodePage = 65001 %>
<%response.expires=0
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
HTProgPrefix="CodeData" %>
<!--#include virtual = "/inc/server.inc" -->
<%
SQL="Select * from CodeMetaDef where codeId=N'" & request.querystring("codeID") & "'"
SET RSCode=conn.execute(SQL)
taskLable="查詢" & HTProgCap

	showHTMLHead()
	formFunction = "query"
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
Sub formSearch()
	reg.action="<%=HTprogPrefix%>DetailList.asp?T=Query&codeID=<%=request.querystring("CodeID")%>&CodeName=<%=request.querystring("CodeName")%>"
	reg.Submit
End Sub

sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm
end sub

sub window_onLoad
	clientInitForm
end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() %>

<!--#include file="CodeDataQueryForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
<title>查詢表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【查詢--代碼ID:<%=request.querystring("CodeID")%>/代碼名稱:<%=session("CodeName")%>】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><a href="Javascript:window.history.back();">回前頁</a>&nbsp;</td>   
	</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=230 valign=top colspan="2">
<% end sub '--- showHTMLHead() ------%>

<% sub ShowHTMLTail() %>  
    </td>
  </tr>  
</table> 
</body>
</html>
<% end sub '--- showHTMLTail() ------%>
