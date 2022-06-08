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
		showDoneBox "資料更新成功！", "0"
		
	elseif xUpForm("submitTask") = "DELETE" then	
	
		sql = "UPDATE kpi_set_score SET xStatus = 'D', delTime = GETDATE(), modTime = GETDATE() WHERE Rank0_2 = '" & xUpForm("deleteId") & "'"
		conn.execute(sql)
		showDoneBox "刪除成功！", "1"
		
	elseif xUpForm("submitTask") = "RECOVERY" then	
	
		sql = "UPDATE kpi_set_score SET xStatus = 'Y', modTime = GETDATE() WHERE Rank0_2 = '" & xUpForm("recoveryId") & "'"
		conn.execute(sql)
		showDoneBox "復原成功！", "1"	
		
	elseif xUpForm("submitTask") = "ADD" then		
	
		'response.write xUpForm("ctNodeIdNew") & "~" & xUpForm("ctNodeGradeNew")
		Dim flag : flag = false
		sql = "SELECT * FROM kpi_set_score WHERE Rank0_0 = '" & xUpForm("ctNodeIdNew") & "' AND Rank0_4 = 'st_22'"
		set rs = conn.execute(sql)
		if not rs.eof then
			flag = true
		end if
		rs.close
		set rs = nothing
		
		if flag then 
			showDoneBox "新增的節點已存在！", "1"
		else
			Dim maxid : maxid = ""
			sql = "SELECT MAX(Rank0_2) AS maxId FROM kpi_set_score WHERE Rank0_4 = 'st_22'"
			set mrs = conn.execute(sql)
			if not mrs.eof then
				maxid = mrs("maxId")
			end if
			set mrs = nothing
			maxId = mid(maxId, len("st_22") + 1)
			maxId = CInt(maxId) + 1
			maxId = "st_22" & maxId
			sql = "INSERT INTO kpi_set_score(Rank0, Rank0_0, Rank0_1, Rank0_2, Rank0_3, Rank0_4, xStatus, addTime, modTime) "
			sql = sql & "VALUES('st_2', '" & xUpForm("ctNodeIdNew") & "', '" & xUpForm("ctNodeGradeNew") & "'," 
			sql = sql & "'" & maxId & "', N'瀏覽行為特殊給點', 'st_22', 'Y', GETDATE(), GETDATE())"
			conn.execute(sql)
			showDoneBox "資料新增成功！", "1"
		end if
					
	else
		showForm()
	end if	

	sub doUpdateDB()
		sql = ""
		for each form in xup.Form					
			if Instr(form.Name, "st_") > 0 then
				if Instr(form.Name, "ctNodeStatus") > 0 then
					sql = sql & "UPDATE kpi_set_score SET xStatus = '" & form & "', modTime = GETDATE() WHERE Rank0_2 = '" & mid(form.Name, len("ctNodeStatus") + 1 ) & "';"
				else					
					sql = sql & "UPDATE kpi_set_score SET Rank0_1 = " & form & ", modTime = GETDATE() WHERE Rank0_2 = '" & form.Name & "';"			
				end if
			end if
		next		
		conn.execute(sql)
	end sub		
	
	Function GetctNodeName(nodeid)
		sql = "SELECT CatName FROM CatTreeNode WHERE CtNodeID = " & nodeid
		set nrs = conn.execute(sql)
		if not nrs.eof then
			GetctNodeName = nrs("CatName")
		end if
		set nrs = nothing
	End Function
%>
<% Sub showDoneBox(lMsg, atype) %>
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
				<% if atype = "1" then%>
					document.location.href = "KpiBrowseSetting.asp?id=st_2&phase=edit"
				<% else %>
					document.location.href = "KpiSettingList.asp"
				<% end if %>
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
		<td width="50%" class="FormLink" align="right"><A href="javascript:window.location.href='/KpiSetting/KpiBrowseSetting.asp" title="回前頁">回前頁</A></td>	
  </tr>
  <tr>
    <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td></tr>
  <tr><td class="Formtext" colspan="2" height="15"></td></tr>
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
		<form method="POST" name="reg" action="KpiBrowseSetting.asp?id=<%=Rank0%>" ENCTYPE="MULTIPART/FORM-DATA">
		<INPUT TYPE="hidden" name="submitTask" value="">
		<INPUT TYPE="hidden" name="deleteId" value="">
		<INPUT TYPE="hidden" name="recoveryId" value="">
		<INPUT TYPE="hidden" name="addId" value="">
		<CENTER>
		<%
			Dim first : first = true
			Dim first2 : first2 = true
			Dim flag : flag = ""
			While not rs.eof 							
				if flag <> rs("Rank0_4") then
					if not first then
						response.write "</TABLE>" & vbcrlf
					else
						first = false
					end if
					response.write "<TABLE width=""100%"" id=""ListTable"">" & vbcrlf
					if rs("Rank0_4") = "st_22" then
						response.write "<TR><Th colspan=""5"">" & rs("Rank0_3") & "</Th></TR>" & vbcrlf
					else
						response.write "<TR><Th colspan=""2"">" & rs("Rank0_3") & "</Th></TR>" & vbcrlf
					end if					
					flag = rs("Rank0_4")
				end if
				
				if rs("Rank0_4") = "st_21" then					
					response.write "<TR>" & vbcrlf
					response.write "<Th width=""30%"">" & rs("Rank0_0") & "：</Th>" & vbcrlf
					if rs("Rank0_2") = "st_211" then
						response.write "<TD class=""eTableContent""><input name=""" & rs("Rank0_2") & """ value=""" & rs("Rank0_1") & """ size=""10"">&nbsp;點&nbsp;/&nbsp;每會員每日</TD>" & vbcrlf				
					else
						response.write "<TD class=""eTableContent""><input name=""" & rs("Rank0_2") & """ value=""" & rs("Rank0_1") & """ size=""10"">&nbsp;點&nbsp;/&nbsp;每頁</TD>" & vbcrlf									
					end if					
					response.write "</TR>" & vbcrlf	
				else
					if first2 then						
            response.write "<TR><th width=""34%"">單元標題</th><th class=""eTableContent"">ctNode值</th>" & _
													"<th class=""eTableContent"">加權點數</th><th class=""eTableContent"">狀態</th><th class=""eTableContent"">&nbsp;</th></TR>" & vbcrlf	
						first2 = false
					end if
						response.write "<TR>" & vbcrlf
            response.write "<td><span class=""eTableContent""><input name=""ctNodeName" & rs("Rank0_0") & """ class=""rdonly"" value=""" & GetctNodeName(rs("Rank0_0")) & """ size=""35""  readonly=""true""></span></td>" & vbcrlf	
            response.write "<TD><input name=""ctNode" & rs("Rank0_0") & """ class=""rdonly"" value=""" & rs("Rank0_0") & """ size=""20"" readonly=""true""></TD>" & vbcrlf	
            response.write "<TD><input name=""" & rs("Rank0_2") & """ value=""" & rs("Rank0_1") & """ size=""10"">&nbsp;點&nbsp;/&nbsp;每頁</TD>" & vbcrlf	
						response.write "<TD><input name=""ctNodeStatus" & rs("Rank0_2") & """ value=""" & rs("xStatus") & """ size=""10"" ></TD>" & vbcrlf	
						if rs("xStatus") = "Y" then
							response.write "<TD><input name=""button"" type=""button"" class=""cbutton"" onClick=""formDelete('" & rs("Rank0_2") & "')"" value =""刪除""></TD>" & vbcrlf	
						else
							response.write "<TD><input name=""button"" type=""button"" class=""cbutton"" onClick=""formRecovery('" & rs("Rank0_2") & "')"" value =""復原""></TD>" & vbcrlf	
						end if
            response.write "</TR>" & vbcrlf						
				end if										
				rs.movenext 
			wend 
			response.write "<TR>" & vbcrlf
      response.write "<td><span class=""eTableContent""><input name=""ctNodeNameNew"" class=""rdonly"" value="""" size=""35""  readonly=""true""></span></td>" & vbcrlf	
      response.write "<TD><input name=""ctNodeIdNew"" value="""" size=""20""></TD>" & vbcrlf	
      response.write "<TD><input name=""ctNodeGradeNew"" value="""" size=""10"">&nbsp;點&nbsp;/&nbsp;每頁</TD>" & vbcrlf	
			response.write "<TD><input name=""ctNodeStatusNew"" value="""" size=""10"" class=""rdonly"" readonly=""true""></TD>" & vbcrlf	
      response.write "<TD><input name=""button"" type=""button"" class=""cbutton"" onClick=""formAdd()"" value =""新增""></TD>" & vbcrlf	
      response.write "</TR>" & vbcrlf						
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
	function formDelete(id) {	
		document.getElementById("submitTask").value = "DELETE";
		document.getElementById("deleteId").value = id;
		document.getElementById("reg").submit();	
	}
	function formRecovery(id) {	
		document.getElementById("submitTask").value = "RECOVERY";
		document.getElementById("recoveryId").value = id;
		document.getElementById("reg").submit();	
	}
	function formAdd() {	
		document.getElementById("submitTask").value = "ADD";		
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
     

