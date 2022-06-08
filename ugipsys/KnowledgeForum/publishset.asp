<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="KForumlist"
HTProgPrefix = ""
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbUtil.inc" -->
<!--#include file = "Activity.inc" -->
<%	
	Dim add : add = request.querystring("add")
	Dim questionId : questionId = request.querystring("questionId")
	Dim id : id = request.querystring("id")	
	Dim nowPage : nowPage = request.querystring("nowPage")
	Dim pageSize : pageSize = request.querystring("pageSize")
	
	if add = "Y" then
		publishArticle
	else
		showForm
	end if
%>
<% 
Sub publishArticle 
	Dim iCount : iCount = 0	
	Dim dCount : dCount = 0
	
	if instr(id, ",") > 0 then
		'---多筆---
		Dim idarr : idarr = split(id, ",")
		for i = 0 to ubound(idarr)
			if idarr(i) <> "" then
				if CheckInsert( idarr(i) ) then
					InsertData( idarr(i) )
					iCount = iCount + 1
				else
					dCount = dCount + 1
				end if
			end if
		next
	else		
		'---單筆---
		if CheckInsert( id ) then
			InsertData( id )
			iCount = iCount + 1
		else
			dCount = dCount + 1
		end if		
	end if
	
	'add by Ivy 2010/5/4 知識家活動 活動問題反餽回主題館,發問者得1分
	
	PublishToSubject questionId
	
	'知識家活動 End
	
	showDoneBox "成功 " & iCount & " 筆, 重複 " & dCount & " 筆", "true"
End Sub
%>
<%
Function CheckInsert( myid )
	Dim subjectId, nodeId, flag
	'---檢查是否有重複insert---
	sql = "SELECT * FROM KnowledgeToSubject WHERE id = " & myid
	set rs = conn.execute(sql)
	if not rs.eof then
		subjectId = rs("subjectId")
		nodeId = rs("ctNodeId")
	end if
	rs.close
	set rs = nothing		
	sql = "SELECT id FROM KnowledgeToSubject WHERE subjectId = " & subjectId & " " & _
				"AND ctNodeId = " & nodeId & " AND iCUItem = " & questionId & " AND status = 'I'"
	set rs = conn.execute(sql)
	if rs.eof then
		flag = true
	else
		flag = false
	end if
	rs.close
	set rs = nothing
	CheckInsert = flag	
End Function
%>
<%
Sub InsertData( myid )
	Dim subjectId, basedsdId, unitId, nodeId, flag
	'---找出此id的相關資料---
	sql = "SELECT * FROM KnowledgeToSubject WHERE id = " & myid
	set rs = conn.execute(sql)
	if not rs.eof then
		subjectId = rs("subjectId")
		basedsdId = rs("baseDSDId")
		unitId = rs("ctUnitId")
		nodeId = rs("ctNodeId")
	end if
	rs.close
	set rs = nothing
	'---找出此問題的資料--
	Dim stitle, xpostdate, xurl
	sql = "SELECT sTitle, CONVERT(varchar, xPostDate, 111) AS xPostDate, topCat " & _
				"FROM CuDTGeneric WHERE iCUItem = " & questionId
	set rs = conn.execute(sql)
	if not rs.eof then
		stitle = rs("sTitle")
		xpostdate = rs("xPostDate")
		xurl = "/Knowledge/Knowledge_cp.aspx?ArticleId=" & questionId & "&ArticleType=A&CategoryId=" & rs("topCat")
	end if
	rs.close
	set rs = nothing
	'---insert into cudtgeneric---	
	sql = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, dEditDate, iDept, xURL, " & _
				"xPostDate, showType, siteId, xNewWindow) VALUES(" & basedsdId & ", " & unitId & ", 'Y', " & pkstr(stitle, "") & ", " & _
				"'hyweb', GETDATE(), '0', " & pkstr(xurl, "") & ", " & pkstr(xpostdate, "") & ", '2', '2', 'Y')"
	sql = "set nocount on;" & sql & "; select @@IDENTITY as NewID"				
	Set rs = conn.Execute(sql)
	xNewIdentity = rs(0)		
	rs.close
	Set rs = nothing
	'---insert into slave table---
	sql = "INSERT INTO CuDTx" & basedsdId & " (gicuitem) VALUES(" & xNewIdentity & ")"
	conn.execute(sql)	
	'---insert into knowledgetosubject---
	sql = "INSERT KnowledgeToSubject(subjectId, baseDSDId, ctUnitId, ctNodeId, status, ICUItem) " & _
				"VALUES(" & subjectId & ", " & basedsdId & ", " & unitId & ", " & nodeId & ", 'I', " & questionId & ")"
	conn.execute(sql)
End Sub
%>
<% Sub showDoneBox(lMsg, btype) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")
				<% if btype = "true" then %>
					document.location.href="cp_question.asp?iCUItem=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>"				
				<% else %>
					history.back
				<% end if %>
			</script>
    </body>
  </html>
<% End sub %>
<% 
Sub showForm 
	
	'---知識家發佈, 資料庫的欄位status = K---
	Dim status : status = "K"	
	
	sql = "SELECT * FROM KnowledgeToSubject WHERE status = " & pkstr(status, "") & " ORDER BY subjectId "
	set rs = conn.execute(sql)	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html> 
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="kpiQuery.asp">
    <title>查詢表單</title>
		<link type="text/css" rel="stylesheet" href="../css/list.css">
		<link type="text/css" rel="stylesheet" href="../css/layout.css">
		<link type="text/css" rel="stylesheet" href="../css/setstyle.css">    
    <style type="text/css">
		<!--
			.style1 {color: #000000}
		-->
    </style>
  </head>
	<body>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
	    <td width="50%" align="left" nowrap class="FormName">發佈設定：知識問題關連主題館單元&nbsp;		</td>
			<td width="50%" class="FormLink" align="right"><A href="cp_question.asp?iCUItem=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>" title="回前頁">回前頁</A></td>	
		</tr>
	  <tr>
	    <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	  </tr>
		<tr>
			<td class="Formtext" colspan="2" height="15"></td>
		</tr>
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
				<form name="iForm" method="post">
				<CENTER>				
        <TABLE width="100%" id="ListTable">
        <TR>
					<th>&nbsp;</th>
          <th>主題館名稱｜ctRoot</th>
					<th>主題單元名稱｜ctUnit值</th>
          <th>節點名稱｜ctNode值</th>
          <th>&nbsp;</th>
        </TR>
				<% while not rs.eof %>
        <TR>
					<td class="eTableContent">
						<input type="checkbox" name="ckbox<%=rs("id")%>" value="<%=rs("id")%>">
					</td>
					<td class="eTableContent">
						<input name="subjectName<%=rs("id")%>" class="rdonly" value="<%=rs("subjectName")%>" size="20"  readonly="true">
            <input name="subjectId<%=rs("id")%>" class="rdonly" value="<%=rs("subjectId")%>" size="10"  readonly="true">
          </td>
					<td class="eTableContent">
						<input name="ctUnitName<%=rs("id")%>" class="rdonly" value="<%=rs("ctUnitName")%>" size="20" readonly="true">
						<input name="ctUnitId<%=rs("id")%>" class="rdonly" value="<%=rs("ctUnitId")%>" size="10" readonly="true">
					</TD>
          <td class="eTableContent">
						<input name="ctNodeName<%=rs("id")%>" class="rdonly" value="<%=rs("ctNodeName")%>" size="20" readonly="true">
						<input name="ctNodeId<%=rs("id")%>" class="rdonly" value="<%=rs("ctNodeId")%>" size="10" readonly="true">
					</TD>
          <td class="eTableContent">
						<input name="button" type="button" class="cbutton" onClick="publishNode(<%=rs("id")%>)" value="發佈">
					</TD>
        </TR>
				<% 
						rs.movenext
					wend 
					rs.close
					set rs = nothing
				%>                         
        </TABLE>
				</CENTER>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td align="center">     
						<input name="button4" type="button" class="cbutton" onClick="formSubmit()" value="整批發佈">
						<input type=button value ="重　設" class="cbutton" onClick="formReset()">
						<input type=button value ="取　消" class="cbutton" onClick="formCancel()">
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
	function formSubmit()
	{		
		var pass = "";
		for( var i = 0; i < document.forms[0].elements.length - 1; i++ ) {
			if( document.forms[0].elements[i].name.substring(0, 5) == "ckbox") {
				if( document.forms[0].elements[i].checked) {
					pass += document.forms[0].elements[i].value + ",";
				}
			}
		}
		if( pass == "") {
			alert("請至少選擇一項");
		}
		else {
			window.location.href = "publishset.asp?add=Y&id=" + pass + "&questionId=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>";			
		}				
	}
	function publishNode(id)
	{		
		window.location.href = "publishset.asp?add=Y&id=" + id + "&questionId=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>";
	}
	function formReset()
	{		
		window.location.href = "publishset.asp?questionId=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>";
	}
	function formCancel()
	{		
		window.location.href = "cp_question.asp?iCUItem=<%=questionId%>&nowPage=<%=nowPage%>&pagesize=<%=pageSize%>";
	}
</script>
<% end sub %>




