<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap = "單元資料維護"
	HTProgFunc = "查詢"
	HTUploadPath = session("Public") + "data/"
	HTProgCode="kpi01"
	HTProgPrefix="" 
	response.charset = "UTF-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
<title></title>
</head>
<body>
	<form method="GET" name="reg" action="thisyear_KpiMangQueryResult.asp">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">本年度KPI整合管理&nbsp;<font size=2>【查詢】</font></td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	<tr>
		<td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
	<tr>
	  <td align="center" colspan="2" width="80%" height="230" valign="top">    
			<CENTER>			
      <TABLE width="100%" id="ListTable">
      <TR>
				<Th width="15%">積分匯出：</Th>
        <TD class="eTableContent">
					<a href="/KpiSetting/thisyear_KpiMemberSummaryExport.asp" target="_blank">產生並匯出會員積分總表</a> (匯出報表前一日統計結果為止)
				</TD>
      </TR>
      </TABLE>
      <TABLE width="100%" id="ListTable">
      <caption>依條件查詢</caption>
			<TR>
				<Th width="15%">帳號：</Th>
				<TD class="eTableContent"><input name="account" size="50"></TD>
			</TR>
			<TR>
				<Th>真實姓名：</Th>
				<TD class="eTableContent"><input name="realname" size="50"></TD>
			</TR>
			<TR>
				<Th>暱稱：</Th>
				<TD class="eTableContent"><input name="nickname" size="50"></TD>
			</TR>
			<TR>
				<Th>會員等級：</Th>
				<TD class="eTableContent">
					<SELECT name="level" size="1">
						<option value="" selected>請選擇</option>
						<option value="1">入門級</option>
						<option value="2">進階級</option>
						<option value="3">高手級</option>
						<option value="4">達人級</option>
					</SELECT>
				</TD>
			</TR>
			<TR>
				<Th>分數區間：</Th>
				<TD class="eTableContent">
					<SELECT name="gradetype" size=1>
						<option value="1" selected>會員總積分</option>
						<option value="2">吸收度(瀏覽行為)積分</option>
						<option value="3">持久度(登入行為)積分</option>
						<option value="4">分享度(互動行為)積分</option>
						<option value="5">貢獻度(內容價值)積分</option>
						<option value="6">踴躍度(活動參與)積分</option>
					</SELECT>
					<br/><br/>
					<input name="gradefrom" size="25">~<input name="gradeto" size="25">
					<br/><br/>
					(可下拉選擇要查詢的積分項目，並請輸入欲查詢之積分區間，如300~500)
				</TD>
			</TR>
			</TABLE>			
			</CENTER>
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center">     
					<input name="QueryBtn" type="submit" class="cbutton" value="查詢">
					<input name="ResetBtn" type="reset" class="cbutton" value="重　填">
				</td>
			</tr>
			</table>
    </td>
  </tr>  
	</table> 
	</form>
</body>
</html>         
<script language="javascript">                   
</script>