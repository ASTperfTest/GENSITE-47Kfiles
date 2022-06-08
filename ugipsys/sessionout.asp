<%@ CodePage = 65001 %><%
	response.redirect "login.aspx?usrid=" & session("userID")
	'response.write "login.aspx?usrid=" & session("userID")
%>