<% Response.Expires = 0 
   HTProgCode = "HTP02"
   'Response.Write Session("UserID")
   uid = Session("UserID")

   dPath = Replace(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/"),"\","\\")
   wPath = Replace(server.MapPath("/site/" & session("mySiteID") & "/Project0516/Public/","\","\\")
   //Response.write "./project0516/transfer.aspx?DBUSERID="& uid & "&dGipDsdPath=" & dPath
   Response.Redirect "./project0516/transfer.aspx?DBUSERID="& uid & "&dGipDsdPath=" & dPath & "&wPublicPath="& wPath
   %>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>管理</title>
</head>
<body >

</body>
</html> 
 
