<% Response.Expires = 0
HTProgCap="課程學員管理"
HTProgFunc="清單"
HTProgCode="PA010"
HTProgPrefix="ppPsnInfo" %>
<!--#INCLUDE FILE="ppPsnInfoListParam.inc" -->
<!--#Include Virtual= "inc/server.inc" -->
<!--#INCLUDE Virtual="inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT htx.psnID, htx.pName, htx.birthDay, htx.eMail, htx.tel, htx.corpName" _
		& " FROM paPsnInfo AS htx" _
		& " WHERE 1=1"
	xpCondition
end if

nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then
         PerPageSize=10
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
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">

		<%if (HTProgRight and 4)=4 then%>
			<A href="ppPsnInfoAdd.asp" title="新增">新增</A>
		<%end if%>
		        <a href="excelgen.asp" title="資料匯出">資料匯出</a>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>
	    </td>
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
	<td class=eTableLable>身份證號</td>
	<td class=eTableLable>姓名</td>
	<td class=eTableLable>出生日</td>
	<td class=eTableLable>eMail</td>
	<td class=eTableLable>連絡電話</td>
	<td class=eTableLable>公司全銜</td>
    </tr>
<%
    for i=1 to PerPageSize
%>
<tr>
<%pKey = ""
pKey = pKey & "&psnID=" & RSreg("psnID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
	<A href="ppPsnInfoEdit.asp?<%=pKey%>">
<%=RSreg("psnID")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
	<A href="ppPsnInfoEdit.asp?<%=pKey%>">
<%=RSreg("pName")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("birthDay")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("eMail")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("tel")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("corpName")%>
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
