<%@ CodePage = 65001 %><%
response.charset = "utf-8"

dim act
dim xPath
dim upPath
dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

ACT = request("submitTask")
xPath =request("xxPath")
upPath = session("uploadPath") & xPath
if right(upPath,1)<>"/" then    
    upPath = upPath & "/"
end if

'response.write upPath & "<HR>"
'response.write ACT & (Act= "刪檔") & "<HR>"


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=response.charset%>">
</head>
<body>
<%


if ACT = "delFiles" Then 
    'response.write "<hr>insertFiles: " & request("insertFiles")
    fList = split(request("insertFiles"),", ")
    for x = 0 to UBound(fList,1)
        'response.write fList(x) & "<HR>"
        fso.DeleteFile(server.MapPath(upPath & fList(x)))
    next
%>
    <script language=javascript>
        alert("刪除 <%=UBound(fList,1)+1%> 個檔案成功");
        //window.navigate "fileMan.asp?xPath=<%=xPath%>"
        window.location.href="fileMan.asp?xPath=<%=xPath%>";
    </script>   
<%  
end if

if ACT = "delDir" Then 
    if right(upPath,1)<>"/" then    upPath = upPath & "/"
    upPath = left(upPath, len(upPath)-1)
    'response.write upPath & "<HR>"
    xPos = inStrRev(upPath, "/")
    if xPos < 1 then    
        response.end
    end if
    
    parentPath = left(upPath, xPos-1)
    fso.DeleteFolder(server.MapPath(upPath))    
%>
    <script language=javascript>
        alert("刪除目錄成功");
        //window.parent.navigate "index.asp"
        //window.navigate "fileMan.asp?xPath=<%=parentPath%>"
        window.parent.Catalogue.location.reload();
    </script>   
<%  
end if


%>
</body>
</html>