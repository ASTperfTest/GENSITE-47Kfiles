end if

nowPage=Request.QueryString("nowPage")  '瞷计

RSreg.Open fSql,Conn,1,1

if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '羆掸计
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=10  
      end if 
      
      RSreg.PageSize=PerPageSize       '–掸计

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '羆计
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2><%=HTProgFunc%></td>
		<td width="50%" class="FormLink" align="right">
