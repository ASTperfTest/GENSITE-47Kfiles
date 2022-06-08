<%@ CodePage = 65001 %>
<%
for each x in request.queryString
	response.write x & "==>" & request(x) & "<br>"
next
%>