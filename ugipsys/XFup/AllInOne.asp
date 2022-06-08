<%@ CodePage = 65001 %>
<%@ Language=VBScript %>
<%
dim i
i=1 
Set fso = server.CreateObject("Scripting.FileSystemObject")


xPath = request("xPath")
upPath = session("uploadPath") & xPath
if right(upPath,1)<>"/" then	upPath = upPath & "/"
if right(xPath,1)<>"/" then	xPath = xPath & "/"
if left(xPath,1)="/" then	xPath = mid(xPath,2)
'response.write upPath
'response.end
Set fldr = fso.GetFolder(server.MapPath(upPath))
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
			<table cellpadding="3" cellspacing="1" border="1">
			        <tr>
					<td  width="100%" colspan="2"  bgcolor="lightblue" align="center">
					路徑：  <%=session("uploadPath")%><%=xpath%>　　
					<%
					response.write "Please choose your upload destination folder!" 
					%>
					</td>
					
				</tr>
			    <tr>
					<th bgcolor="lightblue" align="center">目 錄</th>
					<th bgcolor="lightblue" align="center">檔 案</th>
				</tr>
				<tr>
				<td valign="top"><table width="100%">
			        <tr bgcolor="lightblue">
                                   <td>切換回 <IMG src="ftv2folderclosed.gif">
                                   <A href="AllInOne.asp?xPath=">根目錄</A></td>
                               	</tr>
			        <% 
			           
			            for each sf in fldr.SubFolders
			            
			        %>
			        <tr bgcolor="lightblue">
                                   <td>切換至下層 <IMG src="ftv2folderclosed.gif">
                                   <A href=AllInOne.asp?xPath=<%=xpath%><%=sf.name%>><%=sf.name%></A></td>
                               	</tr>
                               	<%
                               	   i=i+1
                               	   next                               	
                               	 %>
					</table></td>
				<td valign="top" rowspan="2"><table width="100%">
				<tr bgcolor="lightblue">
					<td><INPUT type="button" name="send" VALUE="刪檔"></td>
					<td>檔 名</td>
					<td align="right">大 小</td>
				</tr>	
			        <% 
			            for each sf in fldr.Files
			        %>
			        <tr bgcolor="yellow">
			           <td ><input type="checkbox" name="insertFiles" value="<%=sf.name%>" ></td>
                       <td><%=sf.name%></td>
                        <td align="right"><%=sf.size%></td>
                              	</tr>
                               	<%
                               	    next                               	
                               	 %>
				
				</table></td></tr>
				<tr><td valign="bottom"><table width="100%">
                                <tr><td><input type="hidden" name="xxPath" value="<%=xpath%>">
                                   新增目錄 <input name="newFolder" size=20> <input type="submit" value="新增" border="0" name="send"></td>
                               	</tr>
				</td>
				</tr>	</table>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案1</td>
					<td >
						<input name="photo1" type="file" ID="File4">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案2</td>
					<td >
						<input name="photo2" type="file" ID="File3">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案3</td>
					<td >
						<input name="photo3" type="file" ID="File2">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案4</td>
					<td >
						<input name="photo4" type="file" ID="File1">
					</td>
				</tr>
				<tr>
					<td bgcolor="e1e0ca" align="center">上傳檔案5</td>
					<td >
						<input name="photo5" type="file" >
					</td>
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
