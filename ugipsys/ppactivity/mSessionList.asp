<% Response.Expires = 0
HTProgCap="�ҵ{�޲z"
HTProgFunc="�ҵ{�`��"
HTProgCode="PA005"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#Include Virtual = "inc/server.inc" -->
<!--#INCLUDE Virtual="inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT htx.*, xref1.mValue AS xrStatus " _
		& " , xref2.mValue AS xrActCat, a.actName, a.actTarget" _
		& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS enrollCount" _
		& " FROM (paSession AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mCode = htx.aStatus AND xref1.codeMetaID='classStatus')" _
		& " LEFT JOIN ppAct AS a ON a.actID=htx.actID " _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE 2=2"
	if request.form("htx_coName") <> "" then
		whereCondition = replace("a.actName LIKE '%{0}%'", "{0}", request.form("htx_coName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_coDateS") <> "" then
		rangeS = request.form("htx_coDateS")
		rangeE = request.form("htx_coDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("(htx.bDate BETWEEN '{0}' AND '{1}')", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_actCat") <> "" then
		whereCondition = replace("a.actCat = '{0}'", "{0}", request.form("htx_actCat") )
		fSql = fSql & " AND " & whereCondition
	end if
	fSql = fSql & " ORDER BY actCat"
end if

nowPage=Request.QueryString("nowPage")  '�{�b����

'response.write fSql & "<HR>"
'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then
   totRec=RSreg.Recordcount       '�`����
   if totRec>0 then
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then
         PerPageSize=20
      end if

      RSreg.PageSize=PerPageSize       '�C������

      if cint(nowPage)<1 then
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then
         nowPage=RSreg.PageCount
      end if

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '�`����
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
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>�i�d�ߵ��G�M��j</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>

	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">���]�d��</a>
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
     <font size="2" color="rgb(63,142,186)"> ��
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>��|
        �@<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">��| ���ܲ�
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%>
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>
             <%Next%>
         </select>
         ��</font>
       <% if cint(nowPage) <>1 then %>
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�W�@��</a>
       <%end if%>

        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %>
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�U�@��</a>
        <%end if%>
        | �C������:
       <select id=PerPage size="1" style="color:#FF0000">
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        </select>
     </font>
    <CENTER>
     <TABLE width=95% cellspacing="1" cellpadding="5" class=bg border=0 bgcolor=336699>
     <tr align=center>
	<td class=lightbluetable width=18% >�ҵ{���O</td>
	<td class=lightbluetable width=35%>�ҵ{�W��</td>
	<td class=lightbluetable width=10%>�覸�N�X</td>
	<td class=lightbluetable width=10%>���W�I���</td>
	<td class=lightbluetable width=10%>�H�ƭ���</td>
	<td class=lightbluetable width=8%>���A</td>
	<td class=lightbluetable width=10%>���W�H��</td>

    </tr>
<%
    for i=1 to PerPageSize
%>
<tr>
<%pKey = ""
pKey = pKey & "&paSID=" & RSreg("paSID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("xrActCat")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
	<A href="paSessionEdit.asp?<%=pKey%>" title="<%=RSreg("actTarget")%>">
<%=RSreg("actName")%>
</A>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("paSnum")%>
</font></td>
	<TD class=whitetablebg><font size=2>
	<A href="paSessionEdit.asp?<%=pKey%>" title="<%=RSreg("dtNote")%>">
<%=s7Date(RSreg("bDate"))%>
</A>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSreg("pLimit")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSreg("xrStatus")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
	<A href="psEnrollList.asp?<%=pKey%>"><%=RSreg("enrollCount")%></A>
</font></td>

    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for
    next
   %>
    </TABLE>
    </CENTER>
       <!-- �{������ ---------------------------------->
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
           msgbox "�䤣����, �Э��]�d�߱���!"
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
	    alert("�Х���ܭӮ�!")
	  else
    	select case k
    	end select
	  end if
	end sub

</script>
