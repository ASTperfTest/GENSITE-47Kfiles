<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1

	session.abandon
	Session("pwd") = False
%>
<script language=vbs>
	window.top.location="GipSiteEntry_rdecthinktank.asp?siteID=<%=request("siteID")%>"
</script>
<%
	response.end
%>