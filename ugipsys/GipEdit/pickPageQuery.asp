<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="選取文章"
HTProgFunc="查詢"
HTProgCode="GC1AP1"
HTProgPrefix="pickPage" %>
<!--#include virtual = "/inc/server.inc" -->
    <html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
    <title>查詢表單</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
taskLable="查詢" & HTProgCap

	if request("rtFunc") <>"" then	session("rtFunc") = request("rtFunc")
	if request("wImg") <>"" then	session("wImg") = request("wImg")

	showHTMLHead()
	formFunction = "query"
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm

    end sub

    sub window_onLoad  
	     clientInitForm
    end sub
  </script>	
<%end sub '---- initForm() ----%>

<%Sub showForm() %>
    <script language="vbscript">
    sub formSearchSubmit()
         reg.action="<%=HTprogPrefix%>List.asp"
         reg.Submit
    end Sub
   </script>

<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right">資料範本：</TD>
<TD class="eTableContent"><Select name="htx_ibaseDsd" size=1>
<option value="">請選擇</option>
			<%SQL="Select ibaseDsd,sbaseDsdname from BaseDsd where ibaseDsd IS NOT NULL AND inUse='Y' Order by ibaseDsd"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD></TR>
<TR><TD class="eTableLable" align="right">目錄樹：</TD>
<TD class="eTableContent"><Select name="htx_ctRootId" size=1>
<option value="">--All--</option>
<%
	SQL = "select ctRootId, ctRootName from CatTreeRoot Where 1=1 " _
		& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%'" _
		& " OR '" & session("deptId") & "' LIKE deptId+'%')"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD></TR>
<TR><TD class="eTableLable" align="right">目錄節點名稱：</TD>
<TD class="eTableContent"><input name="htx_CatName" size="40">
</TD></TR>
<TR><TD class="eTableLable" align="right">主題單元名稱：</TD>
<TD class="eTableContent"><input name="htx_CtUnitName" size="40">
</TD></TR>
<TR><TD class="eTableLable" align="right">標題：</TD>
<TD class="eTableContent"><input name="htx_sTitle" size="40">
</TD></TR>
<TR><TD class="eTableLable" align="right">關鍵字：</TD>
<TD class="eTableContent"><input name="htx_xKeyword" size="40">
</TD></TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
 </td></tr>
</table>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	
<%end sub '--- showForm() ------%>

<%sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
			<A href="Javascript:window.close();" title="關閉視窗">關閉視窗</A>
			<A href="Javascript:window.history.back();">回前頁</A> 
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<%end sub '--- showHTMLHead() ------%>

<%sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%end sub '--- showHTMLTail() ------%>
