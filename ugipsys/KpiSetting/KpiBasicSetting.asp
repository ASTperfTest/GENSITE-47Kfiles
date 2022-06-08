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
	if request("submitTask") = "edit" then		
		doUpdateDB()
		showDoneBox("資料更新成功！")		
	else	
		showForm()
	end if	

	sub doUpdateDB()
		
		Dim rank1 : rank1 = request("Rank1")
		Dim rank2 : rank2 = request("Rank2")
		
		sql = "UPDATE kpi_set_ind SET Rank1 = '" & rank1 & "', Rank2 = '" & rank2 & "', modTime = GETDATE() WHERE Rank0 = '" & Rank0 & "'"
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
	
	sql = "SELECT * FROM kpi_set_ind WHERE Rank0 = '" & Rank0 & "'"
	Set rs = conn.execute(sql)
	If rs.eof Then
		response.write "<script>alert('找不到設定值');history.back();</script>"
	Else
		Dim Rank1 : Rank1 = ""
		Dim Rank2 : Rank2 = ""
		if not rs.eof then
			Rank1 = rs("Rank1")
			Rank2 = rs("Rank2")
		end if
		Set rs = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="ePubEdit.asp">
<title>編修表單</title>
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">    
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="50%" class="FormName" align="left"><%=Rank1%>管理&nbsp;<font size=2>【編修】</td>
    <td width="50%" class="FormLink" align="right"><A href="Javascript:window.history.back();" title="回前頁">回前頁</A></td>
	</tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
	  <td class="Formtext" colspan="2" height="15"></td>
	</tr>  
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>        
		<form method="POST" name="reg" action="KpiBasicSetting.asp?id=<%=Rank0%>">
		<INPUT TYPE="hidden" name="submitTask" value="">
		<CENTER>
		<TABLE  id="ListTable">
		<TR>
			<Th><font color="red">*</font>排行榜名稱：</Th>
			<TD class="eTableContent"><input name="Rank1" size="50" value="<%=Rank1%>"></TD>
		</TR>
		<TR>
			<Th>是否公開：</Th>
			<TD class="eTableContent">
				<select name="Rank2" id="select">
					<option value="Y" <% if Rank2 = "Y" Then %> selected <% end if %> >公開</option>
					<option value="N" <% if Rank2 = "N" Then %> selected <% end if %> >不公開</option>
				</select>  
			</TD>
		</TR>
		</TABLE>
		</CENTER>		
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center">     
        <input type="button" value="編修存檔" name="button4"  class="cbutton" onClick="formSubmit()" >         
        <input type="button" value ="重　填" class="cbutton" onClick="resetForm()">
        <input type="button" value ="回前頁" class="cbutton" onClick="javascript:window.location.href='/KpiSetting/kpiSettingList.asp'">
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
		if ( document.getElementById("Rank1").value == "" ) {
			alert("請輸入排行榜名稱");
			document.getElementById("Rank1").focus();			
		}
		else {
			document.getElementById("submitTask").value = "edit";
			document.getElementById("reg").submit();
		}
	}
	function resetForm() {
		window.location.href = "/KpiSetting/KpiBasicSetting.asp?id=<%=Rank0%>";
	}
</script>
<%
	end if 	
End Sub
%>