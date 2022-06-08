<% @ CodePage = 65001 %>
<%
'// purpose: 為了 emily 去 edit Moj 裡 any text of files from server.
'// date: 2006/7/13
'// input:
'//     xPath , source folder.
'//     xFile , source filename.
'//     newFileName , target filename.
'//     filemanxPath , redirect 時, 所需要的參數.

response.charset = "utf-8"
HTProgCode = "GC1AP5"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
    dim FSO
    dim bHaveError
    dim errorMessage
    
    dim xFile
    dim xFileRealpath
    dim filemanxPath
    
    dim renameMode                     '// { "file" | "dir" }
    
    '// initail
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    bHaveError = 0
    errorMessage = ""
    
    '// request
    xPath = request("xPath")
		xpath = "/publish/"
    filemanxPath = request("filemanxPath")      '// 給 filename 用的 path
    xFile = request("xFile")
    renameMode = request("renameMode")
    newFileName = Request("NewFileName")
    xFileRealpath = ""

    If xPath = "" or xfile = "" Or NewFileName = "" Then
        errorMessage = "Error:" & " xPath is empty!"
        bHaveError = 1
    End If
    
    If xfile = "" Then
        errorMessage = "Error:" & " xfile is empty!"
        bHaveError = 1
    End If
    
    If NewFileName = "" Then
        errorMessage = "Error:" & " NewFileName is empty!"
        bHaveError = 1
    End If
    
    if bHaveError <> 1 then
        if not isValidFilename(xFile,errorMessage) then
            bHaveError = 1
        end if
    End If

    if bHaveError <> 1 then
        xFileRealpath = Server.MapPath(xPath & xFile)
        if renameMode <> "dir" then
            isFileFound = FSO.FileExists(xFileRealpath)
        else
            isFileFound = FSO.FolderExists(xFileRealpath)
        end if
    
        If Not isFileFound Then
            errorMessage = "Error:" & xFileRealpath & " does not exist!"
            bHaveError = 1
        end if
    end if
    
    outputMessage = ""
    if bHaveError <> 1 then
        if renameMode <> "dir" then
            FSO.MoveFile xFileRealpath, replace(xFileRealpath , "\" & xfile, "\" & newFileName)
            Msg = "檔案"
        else
            FSO.MoveFolder xFileRealpath, replace(xFileRealpath , "\" & xfile, "\" & newFileName)
            Msg = "資料夾"
        end if
        Msg = Msg & " [ " & xPath & xFile & " ] \n已重新命名為[ " & xPath & newFileName & " ]！"
        outputMessage = Msg
    else
        outputMessage = errorMessage
    End If

%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=<%=response.charset%>">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
<head>
</head>
<body>
<script language="javascript">
<%
if outputMessage <> "" then
    response.write "alert('" & outputMessage & "');" & vbcrlf
end if

if renameMode = "dir" then
    response.write "window.parent.Catalogue.location.reload();" & vbcrlf
end if
%>
    window.location.href="fileMan.asp?xPath=<%=filemanxpath%>";
</script>
</body>
</html>
<%
set FSO = nothing
set conn = nothing

'//================================================================================
'//================================================================================



'// return true: valid file name.
function isValidFilename(byval inputFileName, byRef returnErrorSign)
    dim returnValue
    dim strErrorSign
    dim iPosition

    strErrorSign = "\/:*?""<>|"
    returnValue = true
    iPosition = 0

    if inputFileName <> "" then
        For iPosition = 1 To Len(strErrorSign)
            If Instr(inputFileName,Mid(strErrorSign,iPosition)) > 0 Then
                returnErrorSign =  "檔案名稱不得有" & Mid(strErrorSign,iPosition) & "！"
                returnValue = false
                exit for
            End If
        Next
    end if

    isValidFilename = returnValue
end function

%>