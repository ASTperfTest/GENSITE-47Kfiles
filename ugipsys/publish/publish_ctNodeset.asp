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
HTProgCode="publish"
HTProgPrefix = "publish"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbUtil.inc" -->
<%
	Dim atype : atype = request.querystring("type")
	Dim add : add = request.querystring("add")
	Dim del : del = request.querystring("del")
	Dim id : id = request.querystring("id")
	Dim subjectId : subjectId = request.querystring("subjectId")
	Dim ctNodeId : ctNodeId = request.querystring("ctNodeId")


	if subjectId = "" then subjectId = "0"
	
	if add = "Y" then
		addNode
	elseif del = "Y" then
		delNode
	else
		showForm
	end if
%>
<%


 
Sub addNode 
	Dim flag : flag = true
	Dim status 

    select case atype
        case "knowledge"
            status = "K"
        case "news"
            status = "Y"
        case "subject"
            status = "S"
    end select

	sql = "SELECT id FROM KnowledgeToSubject WHERE subjectId = " & subjectId & " AND ctNodeId = " & ctNodeId & " AND status = " & pkstr(status, "")
	set rs = conn.execute(sql)
	if not rs.eof then
		flag = false
	end if
	rs.close
	set rs = nothing
	if not flag then
		showDoneBox "此節點已存在", "false"
	else		
		sql = "SELECT CatTreeRoot.CtRootID AS subjectId, CatTreeRoot.CtRootName AS subjectName, " & _
					"ISNULL(BaseDSD.iBaseDSD, 0) AS baseDSDId, ISNULL(BaseDSD.sBaseDSDName, ' ') AS baseDSDName, " & _
					"ISNULL(CtUnit.CtUnitID, 0) AS ctUnitId, ISNULL(CtUnit.CtUnitName, ' ') AS ctUnitName, " & _
					"CatTreeNode.CtNodeID AS ctNodeId, CatTreeNode.CatName AS ctNodeName " & _
					"FROM CatTreeRoot INNER JOIN CatTreeNode ON CatTreeRoot.CtRootID = CatTreeNode.CtRootID " & _
					"LEFT OUTER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID " & _
					"LEFT OUTER JOIN BaseDSD ON CtUnit.iBaseDSD = BaseDSD.iBaseDSD " & _
					"WHERE (CatTreeRoot.CtRootID = " & pkstr(subjectId, "") & ") AND (CatTreeNode.CtNodeID = " & pkstr(ctNodeId, "") & ")"
		set rs = conn.execute(sql)
		if not rs.eof then		
			sqlinsert = "INSERT INTO KnowledgeToSubject(subjectId, subjectName, baseDSDId, baseDSDName, ctUnitId, ctUnitName, ctNodeId, ctNodeName, status) " &_
									"VALUES(" & rs("subjectId") & ", " & pkstr(rs("subjectName"), ",") & _
									rs("baseDSDId") & ", " & pkstr(rs("baseDSDName"), ",") & rs("ctUnitId") & ", " & pkstr(rs("ctUnitName"), ",") & _
									rs("ctNodeId") & ", " & pkstr(rs("ctNodeName"), ",") & pkstr(status, "") & ")"		
		end if
		rs.close
		set rs = nothing	
		conn.execute(sqlinsert)
		showDoneBox "新增完成", "true"
	end if
End Sub
%>
<%
Sub delNode
	sql = "UPDATE KnowledgeToSubject SET status = 'Z' WHERE id = " & id
	conn.execute(sql)
	showDoneBox "刪除完成", "true"
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
					document.location.href="publish_ctNodeset.asp?type=<%=atype%>"				
				<% else %>
					history.back
				<% end if %>
			</script>
    </body>
  </html>
<% End sub %>
<% 
Sub showForm 

	Dim status
	if atype = "knowledge" then
		status = "K"	'---知識家發佈, 資料庫的欄位status = K---
	elseif atype = "news" then
		status = "Y"	'---知識家發佈, 資料庫的欄位status = Y---
	elseif atype = "subject" then
		status = "S"	'---知識家發佈, 資料庫的欄位status = S---
	end if
	
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
			<td width="50%" class="FormLink" align="right"><A href="keywords_set.asp" title="回前頁">回前頁</A></td>	
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
          <th>主題館名稱｜ctRoot</th>
					<th>主題單元名稱｜ctUnit值</th>
          <th>節點名稱｜ctNode值</th>
          <th>&nbsp;</th>
        </TR>
		<% while not rs.eof             
        %>
        <TR>
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
						<input name="button" type="button" class="cbutton" onClick="delNode(<%=rs("id")%>)" value="刪除">
					</TD>
        </TR>
				<% 
						rs.movenext
					wend 
					rs.close
					set rs = nothing
				%>       
        <TR>
          <td class="eTableContent">
						<Select name="subjectList" size="1" onchange="subjectListOnChange()">
							<option value="0" <% if subjectId = "" then %>selected<% end if %>>請選擇</option>
							<%
								sql = "SELECT "
								sql = sql & vbcrlf & "	  CtRootID "
								sql = sql & vbcrlf & "	, CtRootName  "
								sql = sql & vbcrlf & "	, (select count(*) from KnowledgeToSubject where status='" & status & "' and KnowledgeToSubject.subjectid = CatTreeRoot.CtRootID) recCount "
								sql = sql & vbcrlf & "FROM CatTreeRoot WHERE (vGroup = 'XX') AND (inUse = 'Y') "
								sql = sql & vbcrlf & "order by recCount "
                                
								set rs = conn.execute(sql)
								while not rs.eof
									temp = ""
									    if trim(subjectId) = trim(rs("CtRootID")) then temp = "selected"									

                                        if rs("recCount") >0 then
                                            response.write "<option value=""" & rs("CtRootID") & """ " & temp & ">" & rs("CtRootName") & " (已設定)</option>" & vbcrlf
                                        else
                                            response.write "<option value=""" & rs("CtRootID") & """ " & temp & ">" & rs("CtRootName") & "</option>" & vbcrlf
                                        end if
                                        									
									rs.movenext
								wend
								rs.close
								set rs = nothing
							%>              
            </select>
          </td>
          <td class="eTableContent">
						<Select name="ctNodeList" size="1">							
							<%
								If subjectId <> "0" then
									sql = "SELECT CatTreeNode.CtNodeID, CatTreeNode.CatName " & _
												"FROM CatTreeRoot INNER JOIN CatTreeNode ON CatTreeRoot.CtRootID = CatTreeNode.CtRootID " & _
												"WHERE (CatTreeRoot.vGroup = 'XX') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') " & _
												"AND (CatTreeRoot.CtRootID = " & pkstr(subjectId, "") & ") AND (CatTreeNode.CtUnitID IS NOT NULL) "
									set rs = conn.execute(sql)
									while not rs.eof
										response.write "<option value=""" & rs("CtNodeID") & """>" & rs("CatName") & "</option>" & vbcrlf
										rs.movenext
									wend
									rs.close
									set rs = nothing
								else
									response.write "<option value=""0"" selected>請先選擇主題館</option>" & vbcrlf
								end if
							%>              
            </select>
          </TD>
          <td class="eTableContent">
						<input name="button" type="button" class="cbutton" value ="新增" onclick="addNode()">
					</TD>
        </TR>
        </TABLE>
				</CENTER>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td align="center">     
						<input name="button4" type="button" class="cbutton" onClick="formCancel()" value="編修存檔">
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
	function subjectListOnChange()
	{
		var subjectId = document.getElementById("subjectList").options[document.getElementById("subjectList").selectedIndex].value;
		window.location.href = "publish_ctNodeset.asp?type=<%=atype%>&subjectId=" + subjectId;
	}
	function addNode()
	{
		var subjectId = document.getElementById("subjectList").options[document.getElementById("subjectList").selectedIndex].value;
		var ctNodeId = document.getElementById("ctNodeList").options[document.getElementById("ctNodeList").selectedIndex].value;		
		window.location.href = "publish_ctNodeset.asp?type=<%=atype%>&add=Y&subjectId=" + subjectId + "&ctNodeId=" + ctNodeId;
	}
	function delNode(id)
	{		
		window.location.href = "publish_ctNodeset.asp?type=<%=atype%>&del=Y&id=" + id;
	}
	function formReset(id)
	{		
		window.location.href = "publish_ctNodeset.asp?type=<%=atype%>";
	}
	function formCancel(id)
	{		
		window.location.href = "keywords_set.asp";
	}
</script>
<% end sub %>




