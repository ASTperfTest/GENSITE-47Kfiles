<%@  codepage="65001" %>
<%
HTProgCap="電子報管理"
HTProgFunc="訂閱清單"
HTProgCode="GW1M51"
HTProgPrefix="epSub" %>
<!--#include virtual = "/inc/server.inc" -->
<%
Response.AddHeader "Content-Disposition","attachment;filename=EmailListing.xls"
Response.ContentType = "application/vnd.ms-excel"
Response.Buffer = False


 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 'fSql=Request.QueryString("strSql")
 '新增查條件
if request("account") <> "" then
	session("account") = request("account")
else 
	session("account") = ""
end if
if request("realname") <> "" then
	session("realname") = request("realname")
else 
	session("realname") = ""
end if
if request("eMail") <> "" then
	session("eMail") = request("eMail")
else 
	session("eMail") = ""
end if

    Dim order 
		
	if request.querystring("order").item = "Non_Member" then ' 會員訂閱清單
	    order  = request.querystring("order").item 
	    fSql = "SELECT htx.email, convert(varchar,htx.createtime,111) createtime, ctRootId, m.account, m.realname"
		fSql = fSql & " from epaper AS htx LEFT JOIN Member AS m ON htx.email=m.email "
	    fSql = fSql & "where  htx.email not in (select email from member where orderepaper ='Y') "
	    fSql = fSql &  " and htx.ctRootId='21'"
	else 
		order  = request.querystring("order").item '非會員訂閱清單
	    fSql = "SELECT htx.email, convert(varchar,htx.createtime,111) createtime, ctRootId, m.account, m.realname"
	    fSql = fSql & ", (SELECT count(*) FROM MemEpaper WHERE MemEpaper.memId=m.account) AS memEPCount "
	    fSql = fSql &  " FROM Epaper AS htx LEFT JOIN Member AS m ON htx.email=m.email"
	    fSql = fSql &  " WHERE htx.ctRootId=" & session("epTreeID")
		
	end if


	' fSql = "SELECT htx.email, convert(varchar,htx.createtime,111) createtime, ctRootId, m.account, m.realname"
	' fSql = fSql & ", (SELECT count(*) FROM MemEpaper WHERE MemEpaper.memId=m.account) AS memEPCount "
	' fSql = fSql &  " FROM Epaper AS htx LEFT JOIN Member AS m ON htx.email=m.email"
	' fSql = fSql &  " WHERE htx.ctRootId=" & session("epTreeID")
	
	
'加入判斷式
	if session("account") <> "" then
		fSql = fSql & " and (m.account is not null and m.account like '%"& session("account") &"%') "
	end if
	if session("realname") <> "" then
		fSql = fSql & " and (m.realname is not null and (m.realname like '%"& session("realname") &"%' or m.realname like '%"& Chg_UNI(session("realname")) &"%' )) "
	end if
	if session("eMail") <> "" then
		fSql = fSql & " and (htx.eMail like '%"& session("eMail") &"%') "
	end if
	fSql = fSql &  " ORDER BY m.account, createtime"

nowPage=Request.QueryString("nowPage")  '現在頁數
'response.write fSql
'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=50  
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      'strSql=server.URLEncode(fSql)
   end if    
end if   

Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">    
    <title></title>
</head>
<body>
<table border="1">
<tr align="left">
<td class="eTableLable">eMail</td>
<td class="eTableLable">訂閱日期</td>
<td class="eTableLable">會員ID</td>
<td class="eTableLable">姓名</td>
</tr>
<%    
do while not RSreg.eof
%>
<tr>
<td><%=trim(RSreg("eMail"))%></td>
<td><%=RSreg("createtime")%></td>
<td><%=RSreg("account")%></td>
<td><%=RSreg("realname")%></td>
</tr>
<%
RSreg.moveNext
loop    
%>
</table>
</body>
</html>
