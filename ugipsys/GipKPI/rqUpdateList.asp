<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="績效管理"
HTProgFunc="稽催清單"
HTProgCode="GC1AP9"
HTProgPrefix="kpiQuery" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
dim xcatCount(10)



if fSql="" then
	fSql = "SELECT htx.ctUnitId, htx.ctUnitName, htx.ctunitexpireday " _
		& " , s.userId, u.userName, u.email" _
		& " , (SELECT max(deditDate) FROM CuDtGeneric WHERE ictunit=htx.ctUnitId) AS lastUpdate" _
		& " , n.catName" _
		& " FROM CtUnit AS htx JOIN CatTreeNode AS n ON n.ctUnitId=htx.ctUnitId" _
		& " JOIN CtUserSet AS s ON s.ctNodeId=n.ctNodeId" _
		& " JOIN InfoUser AS u ON u.userId=s.userId" _
		& " WHERE htx.ctunitexpireday > 0" _
		& " AND (SELECT max(deditDate) FROM CuDtGeneric WHERE ictunit=htx.ctUnitId) " _
  		& " < DATEADD(day, -htx.ctunitexpireday, getdate())" _
		& " ORDER BY htx.ctUnitId, n.ctNodeId, s.userId" 
'		& " JOIN InfoUser AS u ON u.userId=s.userId -- AND u.email is not NULL" _
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
	    <input type="button" value="匯出Excel" Name="OpenAll" class="cbutton" OnClick="window.open 'rqUpdateExcel.asp'">
</td>
	  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="95%" colspan="2" valign=top align="center">
        <table border="0" width="70%" cellspacing="1" class="bluetable" cellpadding="2">
          <tr  class="lightbluetable">
            <th>單元名稱</th>
            <th>更新頻率</th>
            <th>最後更新</th>
            <th>節點項目</th>
            <th>負 責 人</th>
          </tr>
<%
	xCtUnit = ""
	while not RS.eof
		xDept=RS("ctUnitId")
		if xCtUnit <> RS("ctUnitId") then
			if xCtUnit <> "" then		
				response.write "</td></tr>"
			end if

			xCtUnit = RS("ctUnitId")
%>
            <tr style="background:#ffffff"><td><%=RS("ctUnitName")%></td>
            <td align="center"><%=RS("ctunitexpireday")%> 天
            <td align="center"><%=RS("lastUpdate")%>
            <td><%=RS("catName")%>
            <td>
<%
		end if
%>
		‧<A href="rqUpdateMail.asp?user=<%=RS("userId")%>&unit=<%=xCtUnit%>"><%=RS("userName")%></A>
<%
		RS.moveNext
	wend    
 
%>       
		</td>
            </tr>
		</table>
          <script language=vbs>
'          	document.all("CtCount0").innerText = "<%=catCount%>"
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
