<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap="單元資料維護"
	HTProgFunc="編修"
	HTUploadPath=session("Public")+"data/"
	HTProgCode="CW02"
	HTProgPrefix="CW02" 
	
	Dim Rank0 : Rank0 = request.querystring("id")
%>
<!--#include virtual = "/inc/server.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<%
	apath = server.mappath(HTUploadPath) & "\"
	
	if request.querystring("phase") = "edit" then
		Set xup = Server.CreateObject("TABS.Upload")
	else
		Set xup = Server.CreateObject("TABS.Upload")
		xup.codepage=65001
		xup.Start apath
	end if
	
	function xUpForm(xvar)
		xUpForm = xup.form(xvar)
	end function
	
	if xUpForm("submitTask") = "UPDATE" then
		doUpdateDB()
		showDoneBox("資料更新成功！")		
	else	
		showForm()
	end if	

	sub doUpdateDB()
		sql = ""
		for each form in xup.Form		
			if form.Name <> "submitTask" then
				sql = sql & "UPDATE kpi_set_score SET Rank0_1 = " & form & " WHERE Rank0_2 = '" & form.Name & "';"			
			end if
		next
		conn.execute(sql)
	end sub		
%>
<% Sub showDoneBox(lMsg) %>
  <html>
    <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="KpiBasicSetting.asp">
     <title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
   	    alert("<%=lMsg%>")
	    	document.location.href = "KpiSettingList.asp"
			</script>
    </body>
  </html>
<% End sub %> 
<%	
sub showForm()
	
	dim title
	sql = "SELECT Rank1 FROM kpi_set_ind WHERE Rank0 = '" & Rank0 & "'"
	set rrs = conn.execute(sql)
	if not rrs.eof then
		title = rrs("Rank1")
	end if
	rrs.close
	set rrs = nothing
	
	sql = "SELECT * FROM kpi_set_score WHERE Rank0 = '" & Rank0 & "' ORDER BY Rank0_4, Rank0_2"
	Set rs = conn.execute(sql)
	If rs.eof Then
		response.write "<script>alert('找不到設定值');history.back();</script>"
	Else
	
%>
<html> 
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
  <meta name="GENERATOR" content="Hometown Code Generator 1.0">
  <meta name="ProgId" content="kpiQuery.asp">
  <title>查詢表單</title>
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">    
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">	
  <tr>
		<td width="50%" align="left" nowrap class="FormName">會員KPI記錄分數設定－<%=title%>&nbsp;</td>
		<td width="50%" class="FormLink" align="right"><A href="javascript:window.location.href='/KpiSetting/kpiSettingList.asp" title="回前頁">回前頁</A></td>	
  </tr>
  <tr>
    <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td></tr>
  <tr><td class="Formtext" colspan="2" height="15"></td></tr>
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
		<form method="POST" name="reg" action="KpiAdditionalSetting.asp?id=<%=Rank0%>" ENCTYPE="MULTIPART/FORM-DATA">
		<INPUT TYPE="hidden" name="submitTask" value="">
		<CENTER>
		<%
			Dim first : first = true
			Dim flag : flag = ""
			While not rs.eof 							
				if flag <> rs("Rank0_4") then
					if not first then
						response.write "</TABLE>"
					else
						first = false
					end if
					response.write "<TABLE width=""100%"" id=""ListTable"">" & vbcrlf
					response.write "<TR><Th colspan=""2"">" & rs("Rank0_3") & "</Th></TR>" & vbcrlf
					flag = rs("Rank0_4")
				end if
					response.write "<TR>" & vbcrlf
					response.write "<Th width=""30%"">" & rs("Rank0_0") & "：</Th>" & vbcrlf
					response.write "<TD class=""eTableContent""><input name=""" & rs("Rank0_2") & """ value=""" & rs("Rank0_1") & """ size=""10"">&nbsp;點</TD>" & vbcrlf
					response.write "</TR>" & vbcrlf			
				
				rs.movenext 
			wend 
			response.write "</TABLE>" & vbcrlf
		%>		
		</CENTER>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center">     
        <input name="button4" type="button" class="cbutton" onClick="formSubmit()" value="編修存檔">
        <input type=button value ="重　設" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="javascript:window.location.href='/KpiSetting/kpiSettingList.asp'">
			</td>
		</tr>
		</table>   
		</form>
	</td>
</tr>  
</table> 
</body>
</html>
<script language="javascript">
	function formSubmit() {
	
			document.getElementById("submitTask").value = "UPDATE";
			document.getElementById("reg").submit();
	
	}
	function resetForm() {
		window.location.href = "/KpiSetting/KpiBasicSetting.asp?id=<%=Rank0%>";
	}
</script>
<%
	end if 
	Set rs = nothing		
End Sub
%>
     

