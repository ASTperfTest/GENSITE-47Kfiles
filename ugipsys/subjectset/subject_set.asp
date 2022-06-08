<%@ CodePage = 65001 %>
<% Response.Expires = 0 
response.charset = "utf-8"
HTProgCode = "webgeb1"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
id = request("id")
sql ="SELECT NodeInfo.owner,NodeInfo.subject_user,CatTreeRoot.CtRootName,InfoUser.UserName,Dept.deptName FROM "
sql = sql & " Dept INNER JOIN InfoUser ON Dept.deptID = InfoUser.deptID "
sql = sql & " RIGHT OUTER JOIN NodeInfo ON InfoUser.UserID = NodeInfo.owner "
sql = sql & " LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE NodeInfo.CtrootID ='"& id &"'"
set rs = conn.execute(sql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>主題館權限管理</title>
</head>
<body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2">維護專家設定頁</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="CanlendarTarget" value="">    
          <center>    
          <table border="0" width="90%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td class="lightbluetable" align="center">主題館名稱</td>    
			  <td class="lightbluetable" align="center">權限設定</td>
			  <td class="lightbluetable" align="center">維護專家</td>
            </tr>
            <tr>          	  
              <td class="whitetablebg" rowspan=2><%=rs("CtRootName")%></td>
			  <td class="whitetablebg">管理權限 <input type = "button" value="設定" onclick="javascript:location.href='UserQuery.asp?id=<%=id%>&type=1'"></td>
			  <td class="whitetablebg"><%=rs("UserName")%> / <%=rs("deptName")%></td>
            </tr>
			<tr>
				<td class="whitetablebg">上稿權限 <input type = "button" value="設定" onclick="javascript:location.href='UserQuery.asp?id=<%=id%>&type=2'"></td>
				<td class="whitetablebg">
				<%
				if rs("subject_user") <> "" then
					users=split(rs("subject_user"),",")
					for i=0 to Ubound(users)
						sql_ch = "select InfoUser.UserName,Dept.deptName,InfoUser.deptID from "
						sql_ch = sql_ch &" InfoUser INNER JOIN Dept ON InfoUser.deptID = Dept.deptID "
						sql_ch = sql_ch &" where InfoUser.UserID = '"& trim(users(i))&"'"
						'response.write sql_ch
						set rs_ch = conn.execute(sql_ch)
						if not rs_ch.eof then
							response.write rs_ch("UserName") &" / " & rs_ch("deptName") & "&nbsp;"
							response.write "<a href = 'UserListSave.asp?id="&id&"&ckbox="& trim(users(i))&"&type=3&deptID="& rs_ch("deptID") &"'>刪除</a><br>"
						end if
					next
				end if
				%>
				</td>
			</tr>
          </table>
          </center>
      </form>          
    </td>
  </tr>  
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      </td>                                         
  </tr> 
</table>
</body>
</html> 

