<%
    session("ODBCDSN")="driver={SQL Server};server=61.13.76.20;UID=htEMB;PWD=xxemb;DATABASE=dbEMB"
	response.write session("ODBCDSN") & "<HR>"
%>
