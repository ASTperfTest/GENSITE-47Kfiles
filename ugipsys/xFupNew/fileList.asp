<%@ CodePage = 65001 %>
<%
set fso = server.CreateObject("Scripting.FileSystemObject")

xPath = request("path")
set fldr = fso.GetFolder(xPath)
%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<head>
	</head>
	<body>
	
		<form action="/EKP/ekm/manage_doc/doc_batch_insert.jsp" method="post"  ID="Form1">
			<input type="hidden" name="rootPath" value="<%=xpath%>">
			<input type="hidden" name="redirect_url" value="/rdec/aspx/XFup/fsDone.aspx">
			<table cellpadding="3" cellspacing="1" border="0" >
			        <tr>
					<td  width="70%" colspan="5"  bgcolor="lightblue" align="center">
					Folder: <%=xpath%>　　
					</td>
					
				</tr>
				<tr bgcolor="lightblue">
					<td></td>
					<td>檔 名</td>
					<td align="right">大 小</td>
					<td>類 型</td>
					<td>修改日期</td>
				</tr>	
			        <% 
			           
			            for each sf in fldr.Files
			            
			        %>
			        <tr bgcolor="yellow">
			           <td ><input type="checkbox" name="insertFiles" value="<%=sf.name%>" ></td>
                       <td><%=sf.name%></td>
                        <td align="right"><%=sf.size%></td>
                        <td><%=sf.type%></td>
                       <td><%=sf.DateLastModified%></td>
                              	</tr>
                               	<%
                               	    next                               	
                               	 %>
				
				<tr>
					<td colspan="2" align="center">
						<input type="reset" value="清除" border="0"  name="reset" ID="Reset1"> <input type="submit" value="送出" border="0" name="send" id="send">
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
