<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->

<!--#include virtual = "/inc/client.inc" -->
<%
if  request("iCUItem")<>"" then 

sql="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("iCUItem")&"'"

Set rs = conn.Execute(sql)
sql1="SELECT  CuDTGeneric.sTitle as sTitle, CuDTGeneric.fCTUPublic as fCTUPublic, KnowledgeJigsaw.orderSiteUnit as orderSiteUnit, KnowledgeJigsaw.orderSubject as orderSubject, KnowledgeJigsaw.orderKnowledgeTank as orderKnowledgeTank, KnowledgeJigsaw.orderKnowledgeHome as orderKnowledgeHome, KnowledgeJigsaw.parentIcuitem as parentIcuitem , KnowledgeJigsaw.gicuitem as gicuitem FROM  KnowledgeJigsaw INNER JOIN CuDTGeneric ON KnowledgeJigsaw.gicuitem = CuDTGeneric.iCUItem where KnowledgeJigsaw.parentIcuitem='"&request("iCUItem")&"'"
Set rs1 = conn.Execute(sql1)

session("jigsql") = ""
session("jigcheck")=""
end if 
sql2="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("gicuitem")&"'"
Set rs2 = conn.Execute(sql2)
%>
<html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="kpiQuery.asp">
    <title>查詢表單</title>
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
    <script language="javascript" src="js/SS_popup.js" type="text/javascript"></script>
	<style type="text/css">
<!--
.style1 {color: #000000}
-->
    </style>
</head>
<body>
	<form name="iForm" method="post" action="subject_setList.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request("gicuitem")%>">
	<input name="clear" type="hidden" value="1"/>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">
			農業推薦單元知識拼圖內容管理&nbsp;
			<font size=2>【內容條例清單--<%=rs(0)%>主題專區 <%=rs2(0)%>】</font>
		</td>
		<td width="50%" class="FormLink" align="right">
			<A href="/jigsaw/subjectPubList.asp?iCUItem=<%=request.querystring("iCUItem")%>" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
			<CENTER>
			<TABLE width="100%" id="ListTable">
			<TR>
			<% if request("latest") = "1" then  %>
				<Th>主題單元來源：</Th>
				<TD class="eTableContent">
				<SELECT name="CtRootId" size=1>
				<% 					
						sql = "SELECT * FROM CodeMain WHERE codeMetaID = 'jigsaw' ORDER BY mSortValue"
						set rs = conn.execute(sql)
						while not rs.eof 
							response.write "<option value=""" & rs("mCode") & """>" & rs("mValue") & "</option>"
							rs.movenext
						wend
						rs.close
						set rs = nothing
				%>
				</SELECT>
				</TD>
			<% else %>
				<Th>資料庫來源：</Th>
				<TD class="eTableContent">
					<SELECT name="CtRootId" size=1>
						<option value="1" selected>入口網</option>
						<option value="2">主題館</option>
						<option value="3">知識家</option>
						<option value="4">知識庫</option>
					</SELECT>
				</TD>
			<% end if %>
			</TR>
			<TR>
				<Th>節點名稱：</Th>
				<TD class="eTableContent"><input name="CtNodeName" size="30">(輸入節點名稱)</TD>
			</TR>
			<TR>
				<Th>資料標題：</Th>
				<TD class="eTableContent"><label><input name="sTitle" size="30">(輸入資料標題)</label></TD>
			</TR>
			<TR>
				<Th>日期範圍：</Th>
				<TD class="eTableContent">
					<input type="text" name="value(startDate)" id="startDate" size="10" value="">
					<img src="icon_cal.gif" alt="日曆" style="cursor: hand" border="0"
							 onclick="fPopUpCalendarDlgFormat(document.forms[0]['value(startDate)'],0);" />
					~
					<input type="text" name="value(endDate)" size="10" value="">
					<img src="icon_cal.gif"
							 alt="日曆" style="cursor: hand" border="0"
							 onclick="fPopUpCalendarDlgFormat(document.forms[0]['value(endDate)'],0);" />		
					(選擇資料日期區間)
				</TD>
			</TR>
			<TR>
				<Th>狀態：</Th>
				<TD class="eTableContent"><span class="FormLink">
					<Select name="Status" size=1>
						<option value="">請選擇</option>
						<option value="Y" selected>公開</option>
						<option value="N">不公開</option>
					</select>
					<span class="style1">(預設為公開)</span></span>
				</TD>
			</TR>
			</TABLE>
			</CENTER>
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center">     
					<input name="button4" type="submit" class="cbutton" value="查詢">
					<input type="reset" value ="重　填" class="cbutton" >
				</td>	
			</tr>
			</table>
    </td>
  </tr>  
	</table> 
	</form>
</body>
</html>


