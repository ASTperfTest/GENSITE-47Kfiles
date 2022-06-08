<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="電子報發行清單"
HTProgCode="GW1M51"
HTProgPrefix="ePub"
%>
<!--#INCLUDE FILE="ePubListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
	'---vincent---for 2008 epaper---
	Dim eflag
	'If Request("epTreeID") = "143" Then
	If Request("epTreeID") = "21" Then
		eflag = true
	Else
		eflag = false
	End If
	'-------------------------------
	
 	Set RSreg = Server.CreateObject("ADODB.RecordSet")
 	fSql = Request.QueryString("strSql")
 	response.Write request.querystring("strSql") & "<HR>"
 	if request.querystring("epTreeID") <> "" then session("epTreeID") = request.querystring("epTreeID")
 	if request.querystring("ePaperName") <> "" then session("ePaperName") = request.querystring("ePaperName")
	if fSql = "" then
		fSql = "SELECT htx.epubId, htx.pubDate, htx.title, htx.dbDate, htx.deDate, htx.maxNo" _
			& ", (SELECT count(*) FROM EpSend WHERE EpSend.epubId=htx.epubId and EpSend.CtRootID = htx.CtRootID ) AS sendCount " _
			& " FROM EpPub AS htx" _
			& " WHERE htx.ctRootId=" & session("epTreeID")
		xpCondition
	end if

	nowPage = Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


	if Not RSreg.eof then 
   totRec = RSreg.Recordcount       '總筆數
   if totRec > 0 then 
      PerPageSize = cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize = 10  
      end if 
      
      RSreg.PageSize = PerPageSize       '每頁筆數

      if cint(nowPage) < 1 then 
         nowPage = 1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage = nowPage
      totPage = RSreg.PageCount       '總頁數
      strSql = server.URLEncode(fSql)
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
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>--<%=session("ePaperName")%>】</td>
			<td width="50%" class="FormLink" align="right">
				<%if (HTProgRight and 4)=4 then%>
				<A href="ePubAdd.asp?phase=edit&epTreeID=<%=Request("epTreeID")%>" title="新增">新增</A>
				<%end if%>
				<A href="epSubList.asp?epTreeID=<%=session("epTreeID")%>" title="訂閱清單">訂閱清單</A>
				<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr><td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td></tr>
	  <tr><td class="Formtext" colspan="2" height="15"></td></tr>  
  	<tr>
    	<td width="95%" colspan="2" height=230 valign=top>
 			<Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
  			<p align="center">  
				<%If not RSreg.eof then%>     
     			<font size="2" color="rgb(63,142,186)"> 第
     			<font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        	共<font size="2" color="#FF0000"><%=totRec%></font>
       		<font size="2" color="rgb(63,142,186)">筆| 跳至第       
         	<select id=GoPage size="1" style="color:#FF0000">
          	<%For iPage=1 to totPage%> 
            	<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
            <%Next%>   
         	</select>      
         	頁</font>           
       		<% if cint(nowPage) <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       		<%end if%>      
       
        	<% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        	<%end if%>     
        		| 每頁筆數:
       		<select id=PerPage size="1" style="color:#FF0000">            
          	<option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
            <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
            <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
            <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        	</select>     
     		</font>     
    		<CENTER>
     		<TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     		<tr align=left>    
					<td class=eTableLable>發行日期</td>
					<td class=eTableLable>標題</td>
					<td class=eTableLable>資料範圍起日</td>
					<td class=eTableLable>資料範圍迄日</td>
					<td class=eTableLable>資料則數</td>
					<td class=eTableLable>發送份數</td>
					<% If eflag = true Then %>
						<td class=eTableLable>最新農業知識文章</td>
					<% End If %>
    		</tr>	                                
				<% for i=1 to PerPageSize %>                  
				<tr>                  
					<%
						pKey = ""
						pKey = pKey & "&epubid=" & RSreg("epubId") & "&epTreeID=" & Request("epTreeID")
						if pKey<>"" then  pKey = mid(pKey,2)%>
						<TD class=eTableContent><font size=2>
						<A href="ePubEdit.asp?<%=pKey%>&phase=edit"><%=RSreg("pubDate")%></A></font></td>
						<TD class=eTableContent><font size=2><%=RSreg("title")%></font></td>
						<TD class=eTableContent><font size=2><%=RSreg("dbDate")%></font></td>
						<TD class=eTableContent><font size=2><%=RSreg("deDate")%></font></td>
						<TD class=eTableContent><font size=2><%=RSreg("maxNo")%></font></td>
						<TD class=eTableContent><font size=2><A href="eSendList.asp?<%=pKey%>"><%=RSreg("sendCount")%></A>
						<% If eflag = true Then %>
							<TD class=eTableContent><input name="button4" type="button" class="cbutton" onClick="javascript:window.location.href='epaper_setList.asp?<%=pKey%>'" value="設定"></td>
						<% End If %>
						</font></td>
    		</tr>
    		<%
        	 	RSreg.moveNext
	         	if RSreg.eof then exit for 
    			next 
  			%>
    		</TABLE>
    		</CENTER>
       	<!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr><td width="100%" colspan="2" align="center"></td></tr>
	</table> 
	</form>
	</body>
</html>                                 
<% else %>
	<script language=vbs>
  	msgbox "找不到資料, 請重設查詢條件!"
		'window.history.back
  </script>
<% end if %>

<script Language=VBScript>
  dim gpKey
  sub GoPage_OnChange      
  	newPage=reg.GoPage.value     
    document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
  end sub      
     
  sub PerPage_OnChange                
  	newPerPage=reg.PerPage.value
    document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
  end sub 

  sub setpKey(xv)
   	gpKey = xv
  end sub

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("請先選擇個案!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>
