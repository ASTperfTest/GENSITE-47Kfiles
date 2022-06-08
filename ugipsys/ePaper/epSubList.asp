<%@  codepage="65001" %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="訂閱清單"
HTProgCode="GW1M51"
HTProgPrefix="epSub" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 
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
		
	if request.querystring("order").item = "Non_Member" then '會員訂閱清單
	    order  = request.querystring("order").item 
	    fSql = "select htx.*, m.account, m.realname "
		fSql = fSql & " from epaper AS htx LEFT JOIN Member AS m ON htx.email=m.email "
	    fSql = fSql & "where  htx.email not in (select email from member where orderepaper ='Y') "
	    fSql = fSql &  " and htx.ctRootId='21'"
	else 
		order  = request.querystring("order").item '非會員訂閱清單
	    fSql = "SELECT htx.*, m.account, m.realname"
	    fSql = fSql & ", (SELECT count(*) FROM MemEpaper WHERE MemEpaper.memId=m.account) AS memEPCount "
	    fSql = fSql &  " FROM Epaper AS htx LEFT JOIN Member AS m ON htx.email=m.email"
	    fSql = fSql &  " WHERE htx.ctRootId=" & session("epTreeID")
		
	end if
	

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
	fSql = fSql &  " ORDER BY m.account, htx.createtime"

nowPage=Request.QueryString("nowPage")  '現在頁數
' response.write fSql
' response.end

' ----------HyWeb GIP DB CONNECTION PATCH----------
' RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)
' ----------HyWeb GIP DB CONNECTION PATCH----------

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
      ' strSql=server.URLEncode(fSql)
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
    <link rel="stylesheet" href="/inc/setstyle.css">
    <title></title>
</head>
<body>
    <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
        <tr>
            <td width="50%" class="FormName" align="left">
                <%=HTProgCap%>
				
				<%if request.querystring("order").item <> "Non_Member" then%>
                    &nbsp;<font size="2">【<%=HTProgFunc%>】</td>
				<%else%>
				    &nbsp;<font size="2">【非會員訂閱清單】</td>
				<%end if%>
				
		    <%if request.querystring("order").item <> "Non_Member" then%>
            <td width="50%" class="FormLink" align="right"><a href="ePubList.asp?epTreeID=<%=Request("epTreeID")%>" title="回電子報條列">回電子報條列</a> </td>
            <%else%>
			<td width="50%" class="FormLink" align="right"><A href="epSubList.asp?epTreeID=<%=session("epTreeID")%>" title="訂閱清單">回訂閱清單</A> </td>
			<%end if%>
		</tr>
        <tr>
            <td width="100%" colspan="2">
                <hr noshade size="1" color="#000080">
            </td>
        </tr>
        <tr>
            <td class="Formtext" colspan="2" height="15"></td>
        </tr>
        <tr>
            <td width="95%" colspan="2" height="230" valign="top">
                <form name="reg" method="POST" action="<%=HTprogPrefix%>List.asp">
                    <p align="center">
                        <%If not RSreg.eof then%>
						
						<%if request.querystring("order").item <> "Non_Member" then%>
                            <font size="2" color="rgb(63,142,186)">第 <font size="2" color="#FF0000">
                                <%=nowPage%>
                                /<%=totPage%></font>頁| 共<font size="2" color="#FF0000"><%=totRec%></font> <font size="2" color="rgb(63,142,186)">筆| 跳至第
							        <select id="GoPage" size="1" style="color: #FF0000">
                                        <%For iPage=1 to totPage%>
                                        <option value="<%=iPage%>" <%if iPage=cint(nowPage) then %>selected<%end if%>>
                                            <%=iPage%>
                                        </option>
                                        <%Next%>
                                    </select>
                                    頁</font>
                                <% if cint(nowPage) <>1 then %>
                                |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>">上一頁</a>
                                <%end if%>
                                <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %>
                                |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>">下一頁</a>
                                <%end if%>
                                | 每頁筆數:
                                <select id="PerPage" size="1" style="color: #FF0000">
                                    <option value="50" <%if PerPageSize=50 then%> selected<%end if%>>50</option>
                                    <option value="100" <%if PerPageSize=100 then%> selected<%end if%>>100</option>
                                    <option value="200" <%if PerPageSize=200 then%> selected<%end if%>>200</option>
                                    <option value="300" <%if PerPageSize=300 then%> selected<%end if%>>300</option>
                                </select>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="epSubListSearch.asp?epTreeID=<%=Request("epTreeID")%>" title="查詢電子報" align="right">查詢電子報</a> </font>
                            <%else%>
							    <font size="2" color="rgb(63,142,186)">第 <font size="2" color="#FF0000">
                                <%=nowPage%>
                                /<%=totPage%></font>頁| 共<font size="2" color="#FF0000"><%=totRec%></font> <font size="2" color="rgb(63,142,186)">筆| 跳至第
							        <select id="GoPage_NonMember" size="1" style="color: #FF0000">
                                        <%For iPage=1 to totPage%>
                                        <option value="<%=iPage%>" <%if iPage=cint(nowPage) then %>selected<%end if%>>
                                            <%=iPage%>
                                        </option>
                                        <%Next%>
                                    </select>
                                    頁</font>
                                <% if cint(nowPage) <>1 then %>
                                |<a href="<%=HTProgPrefix%>List.asp?order=<%=order%>&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>">上一頁</a>
                                <%end if%>
                                <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %>
                                |<a href="<%=HTProgPrefix%>List.asp?order=<%=order%>&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>">下一頁</a>
                                <%end if%>
                                | 每頁筆數:
                                <select id="PerPage_NonMember" size="1" style="color: #FF0000">
                                    <option value="50" <%if PerPageSize=50 then%> selected<%end if%>>50</option>
                                    <option value="100" <%if PerPageSize=100 then%> selected<%end if%>>100</option>
                                    <option value="200" <%if PerPageSize=200 then%> selected<%end if%>>200</option>
                                    <option value="300" <%if PerPageSize=300 then%> selected<%end if%>>300</option>
                                </select>
							<%end if%>

							<%if request.querystring("order").item <> "Non_Member" then%>
							&nbsp;&nbsp;<a href="ExportepSubList.asp">匯出</a>
							<%else%>
							&nbsp;&nbsp;<a href="ExportepSubList.asp?order=Non_Member">匯出</a>
							<%end if%>
                        <center>
                            <table width="100%" cellspacing="1" cellpadding="0" class="bg">
                                <tr align="left">
                                    <td class="eTableLable">eMail</td>
                                    <td class="eTableLable">訂閱日期</td>
                                    <td class="eTableLable">會員ID</td>
                                    <td class="eTableLable">姓名</td>
                                </tr>
                                <%for i=1 to PerPageSize%>
                                <tr>
                                    <%pKey = ""
                                    pKey = pKey & "&eMail=" & RSreg("eMail")
                                    if pKey<>"" then  pKey = mid(pKey,2)%>
                                    <td class="eTableContent"><font size="2"><a href="epSubEdit.asp?<%=pKey%>">
                                        <%=RSreg("eMail")%>
                                    </a></font></td>
                                    <td class="eTableContent"><font size="2">
                                        <%=RSreg("createtime")%>
                                    </font></td>
                                    <td class="eTableContent"><font size="2"><a href="/member/member_edit.asp?account=<%=RSreg("account")%>">
                                        <%=RSreg("account")%>
                                    </a></font></td>
                                    <td class="eTableContent"><font size="2">
                                        <%=RSreg("realname")%>
                                    </font></td>                                    
                                </tr>
                                <% RSreg.moveNext
                                   if RSreg.eof then exit for 
                                        next 
								%>
                            </table>
                        </center>
                        <!-- 程式結束 ---------------------------------->
                </form>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="2" align="center"></td>
        </tr>
    </table>
    </form>
</body>
</html>
<%else%>

<script language="vbs">
           msgbox "找不到資料, 請重設查詢條件!"
	       window.history.back
      </script>

<%end if%>

<script language="VBScript">
  dim gpKey
     sub GoPage_OnChange      
        newPage=reg.GoPage.value     
        document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>"    	
	 end sub  
     
	 sub GoPage_NonMember_OnChange      
	 
        newPage=reg.GoPage_NonMember.value     
        document.location.href="<%=HTProgPrefix%>List.asp?order=<%=order%>&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>"    	
	 end sub 
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>&pagesize=" & newPerPage & "&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>"
     end sub 
	 
	 sub PerPage_NonMember_OnChange                
           newPerPage=reg.PerPage_NonMember.value
           document.location.href="<%=HTProgPrefix%>List.asp?order=<%=order%>&nowPage=<%=nowPage%>&pagesize=" & newPerPage & "&account=<%=session("account")%>&realname=<%=session("realname")%>&eMail=<%=session("eMail")%>"
     end sub 
	 
	 

     sub setpKey(xv)
     	gpKey = xv
     end sub


</script>

