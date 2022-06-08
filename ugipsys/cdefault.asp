<%@ CodePage = 65001 %><%
if 	session("mySiteID") = "" then	response.redirect "default.htm"


Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
If Session("pwd") = True Then
	frameRows = "87,*"
	frameRows = "60,*"
'	frameCols = "159,*"
	frameCols = "180,*"
	frameCols = "167,*"
   topfile="ntopmenu.asp"
   leftfile="nstackMenu.asp"
   leftfile="function.asp"
   rightfile="GipEdit/CtNodeTList.asp?CtRootID=4"      
   if session("ModifyPassword") = True Then
   	rightfile="/user/userUpdate.asp"      
   End if
Else
	frameRows = "0,*"
	frameCols = "0,*"
	topfile=""
   	leftfile=""
   	rightfile="gipLogin.asp"   
End If  
%>
<html>
<head>
<title><%=session("mySiteName")%> 後台管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<frameset rows="<%=frameRows%>" cols="*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="topFrame" scrolling="NO" noresize src="<%=topfile%>" marginwidth="0" marginheight="0" frameborder="NO" >
  <frameset cols="<%=frameCols%>" frameborder="NO" border="0" framespacing="0" rows="*" name="f"> 
    <frame name="leftFrame" noresize scrolling="NO" src="<%=leftfile%>" marginwidth="0" marginheight="0" frameborder="NO">
    <frame name="mainFrame" src="<%=rightfile%>" marginwidth="10" marginheight="11">
  </frameset>
</frameset>
<noframes><body bgcolor="#FFFFFF">

</body></noframes>
</html>









