<% Response.Expires = 0
HTProgCap="課程管理"
HTProgFunc="課程總表"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<% response.expires = 0 %>
<!--#Include Virtual = "inc/server.inc" -->
<!--#Include Virtual="inc/dbutil.inc" -->

<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT htx.*, xref1.mValue AS xrActCat " _
		& ", (SELECT count(*) FROM paSession AS s WHERE s.actID=htx.actID) AS sessionCount" _
		& " FROM (ppAct AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mCode = htx.actCat AND xref1.codeMetaID='ppActCat')" _
		& " WHERE 2=2"
	if request.form("htx_coName") <> "" then
		whereCondition = replace("htx.coName LIKE '%{0}%'", "{0}", request.form("htx_coName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_coDesc") <> "" then
		whereCondition = replace("htx.coDesc LIKE '%{0}%'", "{0}", request.form("htx_coDesc") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_coDateS") <> "" then
		rangeS = request.form("htx_coDateS")
		rangeE = request.form("htx_coDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("((htx.coBDate BETWEEN '{0}' AND '{1}') OR ('{0}' BETWEEN htx.coBDate AND htx.coEDate))", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_ceDateS") <> "" then
		rangeS = request.form("htx_ceDateS")
		rangeE = request.form("htx_ceDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("((htx.ceBDate BETWEEN '{0}' AND '{1}') OR ('{0}' BETWEEN htx.ceBDate AND htx.ceEDate))", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_status") <> "" then
		whereCondition = replace("htx.status = '{0}'", "{0}", request.form("htx_status") )
		fSql = fSql & " AND " & whereCondition
	end if
	for each x in request.form
	    if request(x) <> "" then
            if mid(x,2,3) = "fx_" then
	    	   select case left(x,1)
		          case "s"
		            	fSql = fSql & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
		          case else
		            	fSql = fSql & " AND " & mid(x,5) & " LIKE '%" & request(x) & "%'"
		         end select
	        end if
	    end if
	next
	fSql = fSql & " ORDER BY actCat"
end if

nowPage=Request.QueryString("nowPage")  '現在頁數

'response.write fSql & "<HR>"
'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.EOF then
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then
         PerPageSize=20
      end if

      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then
         nowPage=RSreg.PageCount
      end if

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if
end if
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【查詢結果清單】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTprogPrefix%>Add.asp">新增</a>
	       <%end if%>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">重設查詢</a>
	       <%end if%>
    </td>
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
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
     <TABLE width=90% cellspacing="1" cellpadding="8" class=bg bgcolor=navy>
     <tr align=left bgcolor=white>
	<td align=center class=lightbluetable width=10%>課程代碼</td>
	<td align=center class=lightbluetable width=15%>課程類別</td>
	<td align=center class=lightbluetable>項目名稱</td>
	<td align=center class=lightbluetable>課程對象</td>
	<td align=center class=lightbluetable width=8%>梯次</td>
	<td align=center class=lightbluetable width=8%>報名表</td>
    </tr>
<%
    for i=1 to PerPageSize
%>
<tr bgcolor=white>
<%pKey = ""
pKey = pKey & "&actID=" & RSreg("actID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("actID")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("xrActCat")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
	<A href="paActEdit.asp?<%=pKey%>">
<%=RSreg("actName")%>
</A>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("actTarget")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<A href="paSessionList.asp?<%=pKey%>&x=1"><%=RSreg("sessionCount")%></A>
</font></td>
        <TD class=whitetablebg align=center><font size=2>
<A href="paInfoDesign.asp?<%=pKey%>&x=1">設計</A>
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
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table>
</form>
</body>
</html>
<%else%>
      <script language=vbs>
           msgbox "找不到資料, 請重設查詢條件!"
'	       window.history.back
      </script>
<%end if%>

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
