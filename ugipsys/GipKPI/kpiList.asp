<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="績效管理"
HTProgFunc="查詢清單"
HTProgCode="GC1AP9"
HTProgPrefix="kpiQuery" %>
<!--#INCLUDE FILE="kpiListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
dim xcatCount(10)

 fSql=Request.QueryString("strSql")


	for each x in request.form
'		response.write x & "==>" & request(x) & "<BR>"
	next

if fSql="" then
	fSql = "SELECT idept, count(*) AS xCount, " _
		& " idept, u.deptId," _
		& " min(Dept.abbrName) AS dpName, min(xref1.mvalue) AS xTopCat" _
		& " FROM CuDtGeneric AS htx" _
		& " LEFT JOIN CodeMain AS xref1 ON xref1.mcode = htx.topCat AND xref1.codeMetaId='topDataCat'" _
		& " LEFT JOIN InfoUser AS u ON htx.ieditor=u.userId" _
		& " LEFT JOIN Dept  ON Dept.deptId = htx.idept " _
		& " WHERE 1=1 " 

	xConStr = ""
	xpCondition
	fSql = fSql & " GROUP BY htx.idept, u.deptId" _
				& " ORDER BY htx.idept, u.deptId" 
end if
'	response.write fSql & "<HR>"
	set RS = conn.execute(fSql)
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
	    <td width="75%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="25%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="kpiQuery.asp" title="重設查詢條件">重設查詢</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext"><table border=0><%=xConStr%></table>
	    </td>    
	    <td class="Formtext" align=right>      
	    <input type="button" value="全部關閉" Name="OpenAll" class="cbutton" OnClick="chdisplayall()">
</td>
	  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="95%" colspan="2" valign=top align="center">
       <table border="0" width="70%" class="bluetable" cellspacing="0" cellpadding="3"> 
        <tr style="color:#ffffff"> 
          <td width="70%" nowrap>
           </td>
          <td width="15%" align="right" nowrap>總篇數：</td>
          <td width="15%" align="left" ID="CtCount0" nowrap></td>
        </tr>
      </table>
        <table border="0" width="70%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr style="background:#ffffff">
            <td>機關</td>
            <td align="right">自上</td>
            <td align="right">代上</td>
            <td align="right">合計</td>
          </tr>
<%
	cDept = ""
	catCount = 0
	baseCondition = "kpiDList.asp?" 
	if request.form("htx_CtUnitID") <> "" then _
		baseCondition = baseCondition & "htx_CtUnitID=" & request.form("htx_CtUnitID") & "&"
	if request.form("htx_iBaseDSD") <> "" then _
		baseCondition = baseCondition & "htx_iBaseDSD=" & request.form("htx_iBaseDSD") & "&"
	if request.form("htx_IDateS") <> "" then _
		baseCondition = baseCondition & "htx_IDateS=" & request.form("htx_IDateS") & "&"
	if request.form("htx_IDateE") <> "" then _
		baseCondition = baseCondition & "htx_IDateE=" & request.form("htx_IDateE") & "&"
	if request.form("htx_topCat") <> "" then _
		baseCondition = baseCondition & "htx_topCat=" & request.form("htx_topCat") & "&"
	while not RS.eof
		xDept=RS("idept")
		if xDept <> cDept then
			if cDept <> "" then		
%>
            <tr class="lightbluetable"><td><%=dpName%></td>
            <td align="right"><%=selfCount%>
            <td align="right"><%=otherCount%>
            <td align="right"><A href="<%=baseCondition%>htx_idept=<%=cDept%>"><%=selfCount+otherCount%></A></td>
            </tr>
<%
			end if

			cDept = xDept
			dpName = RS("dpName")
			
			selfCount = 0
			otherCount = 0
		end if
		catCount = catCount + RS("xCount")
		if RS("deptId") = xDept then	
			selfCount = selfCount + RS("xCount")
		else
			otherCount = otherCount + RS("xCount")
		end if
		RS.moveNext
	wend    
 
%>       
            <tr class="lightbluetable"><td><%=dpName%></td>
            <td align="right"><%=selfCount%>
            <td align="right"><%=otherCount%>
            <td align="right"><A href="<%=baseCondition%>htx_idept=<%=cDept%>"><%=selfCount+otherCount%></A></td>
            </tr>
		</table>
          <script language=vbs>
          	document.all("CtCount0").innerText = "<%=catCount%>"
          </script>
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</form>
</body>
</html>                                 

<script Language=VBScript>
<%if cno<>"" then%>
sub chdisplayall()   
IF document.all("OpenAll").value="全部展開" Then   
 for i=1 to <%=cno%>          
  if document.all("M" & i).style.display="none" then         
    document.all("M" & i).style.display=""         
    document.all("I" & i).src="/images/2.gif"         
  end if         
 next    
    document.all("OpenAll").value="全部關閉"    
ElseIF document.all("OpenAll").value="全部關閉" Then   
 for i=1 to <%=cno%>          
  if document.all("M" & i).style.display="" then         
    document.all("M" & i).style.display="none"         
    document.all("I" & i).src="/images/1.gif"         
  end if         
 next    
    document.all("OpenAll").value="全部展開"    
End IF            
end sub   
<%end if%>      
         
sub chdisplay(k)         
  if document.all("M" & k).style.display="" then         
    document.all("M" & k).style.display="none"     
    document.all("I" & K).src="/images/1.gif"         
  else         
    document.all("M" & k).style.display=""         
    document.all("I" & K).src="/images/2.gif"         
  end if         
end sub         
     

</script>
