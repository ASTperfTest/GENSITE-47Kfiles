<%@ CodePage = 65001 %>
<%@ Language=VBScript %>
<%
dim i
i=1 
Set fso = server.CreateObject("Scripting.FileSystemObject")


xPath = request("xPath")
if xPath = "" then xPath = "/public"
if right(xPath,1)<>"/" then	xPath = xPath & "/"
Set fldr = fso.GetFolder(server.MapPath(xPath))
'Set fldr = fso.GetFolder("public")
'response.write "Folder name is: " &  fldr.Name & "<HR>"
'for each sf in fldr.SubFolders
'	response.write sf.name & "<BR>"
'next

 
	%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<head>
	</head>
	<body>
	
		<form action="upload.asp" method="post" enctype="multipart/form-data" ID="Form1">
			<table cellpadding="3" width="680" cellspacing="1" border="0" >
			        <tr>
					<td  width="70%" colspan="5"  bgcolor="lightblue" align="center">
					Root: <%=xpath%>　　
					<%
					response.write "Please choose your upload destination folder!" 
					%>
					</td>
					
				</tr>
			        <tr width="10%" bgcolor="yellow" align="center">
			           <td ><input type=radio name=uc value="<%=xpath%>"></td>
                                   <td><font size="4">/</A></font></td>
                               	</tr>
			        <% 
			           
			            for each sf in fldr.SubFolders
			            
			        %>
			        <tr width="10%" bgcolor="yellow" align="center">
			           <td ><input type=radio name=uc value=<%=xpath%><% =sf.name %> ></td>
                                   <td><font size="4"><A href=index.asp?xPath=<%=xpath%><%=sf.name%>><%=sf.name%></A></font></td>
                               	</tr>
                               	<%
                               	   i=i+1
                               	   next                               	
                               	 %>
				<tr>
					<td  width="20%" bgcolor="e1e0ca" align="center">檔案上傳人名字</td>
					<td  colspan="4">
						<input type="text" size="40" name="subject" maxlength="50" ID="Text1">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案1</td>
					<td >
						<input name="photo1" type="file" ID="File4">
					</td>
					<td bgcolor="e1e0ca" align="center">檔案內容簡述</td>
					<td ><input type="text" size="20" name="desc1" maxlength="50" ID="Text21"></td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案2</td>
					<td >
						<input name="photo2" type="file" ID="File3">
					</td>
					<td bgcolor="e1e0ca" align="center">檔案內容簡述</td>
					<td ><input type="text" size="20" name="desc2" maxlength="50" ID="Text21"></td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案3</td>
					<td >
						<input name="photo3" type="file" ID="File2">
					</td>
					<td bgcolor="e1e0ca" align="center">檔案內容簡述</td>
					<td ><input type="text" size="20" name="desc3" maxlength="50" ID="Text21"></td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案4</td>
					<td >
						<input name="photo4" type="file" ID="File1">
					</td>
					<td bgcolor="e1e0ca" align="center">檔案內容簡述</td>
					<td ><input type="text" size="20" name="desc4" maxlength="50" ID="Text21"></td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案5</td>
					<td >
						<input name="photo5" type="file" >
					</td>
					<td bgcolor="e1e0ca" align="center">檔案內容簡述</td>
					<td ><input type="text" size="20" name="desc5" maxlength="50" ID="Text21"></td>
				</tr>
				
				<tr>
					<td colspan="2" align="center">
						<input type="reset" value="清除" border="0"  name="reset" ID="Reset1"> <input type="submit" value="送出" border="0" name="send" id="send">
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
