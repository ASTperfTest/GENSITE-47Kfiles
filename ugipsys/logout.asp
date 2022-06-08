<%@ CodePage = 65001 %><%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1

session.abandon
Session("pwd") = False
Session("siteID") = False
%>
<script language=vbs>
	//window.top.location="GipSiteEntry.asp?siteID=<%=request("siteID")%>"
	window.top.location="doorlist.htm"
</script>
<%
response.end
%>