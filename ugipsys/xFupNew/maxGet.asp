<%@ CodePage = 65001 %><%
'// purpose: ¬°¤F download any type of files from server.
'// date: 2006/7/13

Response.Buffer = True
response.charset="utf-8"
Response.Clear

HTProgCode = "GC1AP5"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
    dim xFile
    dim sourceFile
    dim myStream
    dim FSO
    dim myFile

    dim strMessage              '// message
    dim bHaveError              '// is error flag.


    '// request
    xFile = request("xFile")

    '// initial
    bHaveError = Fasle
    strMessage = ""
    sourceFile = Server.MapPath(xFile)

    Set myStream = Server.CreateObject("ADODB.Stream")
    myStream.Open
    myStream.Type = 1

    Set FSO = Server.CreateObject("Scripting.FileSystemObject")


    '// processing

    '// step 1: check file exist.
    if isFileExist(sourceFile) = 0 then
        strMessage = "<h1>Error:</h1>" & sourceFile & " does not exist<p>"
        bHaveError = True
    end if


    if not bHaveError then
        '// step 2: get file info.
        Set myFile = fso.GetFile(sourceFile)
        intFilelength = myFile.size
        myStream.LoadFromFile(sourceFile)

        Response.AddHeader "Content-Disposition", "attachment; filename=" & myFile.name
        Response.AddHeader "Content-Length", intFilelength
        Response.CharSet = response.charset
        Response.ContentType = "application/octet-stream"
        Response.BinaryWrite myStream.Read
        Response.Flush
    else
        response.write strMessage
    end if

Set myStream = Nothing
set myFile = nothing
Set FSO = Nothing

'//================================================================================
'//================================================================================


'// purpose: check a file is exist.
'// return: 0: not exist.
'// return: 1: exist.
'// ex: ret = isFileExist(inputFilepath)
function isFileExist(byval inputFilepath)
    dim returnValue
    dim FSOCheck

    returnValue = 0
    Set FSOCheck = Server.CreateObject("Scripting.FileSystemObject")
    If FSOCheck.FileExists(inputFilepath) Then
        returnValue = 1
    end if

    isFileExist = returnValue
end function
%>