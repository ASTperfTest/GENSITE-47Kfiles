<%@ CodePage = 65001 %><%
'// purpose: bypass firewall check, decode file.
'// date: 2006/7/13
'// input:
'//     xPath , source folder.
'//     xFile , source filename.
'//     filemanxPath , redirect 時, 所需要的參數.

response.charset="utf-8"
%>
<!--#include virtual = "/inc/client.inc" -->
<%
    const myPrivateKey = 18     '// decode private key value.
    dim xFile
    dim xPath
    dim filemanxPath
    dim sourceFile
    dim FSO                     '// check file exist object.
    dim myStream                '// read file stream
    dim myFile                  '// get file info
    dim intFilelength
    dim outputByte              '// each byte to output
    dim totalOutputByte         '// total byes to output
    dim outputMessage              '// message
    dim bHaveError              '// is error flag.


    '// request
    xFile = request("xFile")
    xPath = request("xPath")
    filemanxPath = request("filemanxPath")      '// 給 filename 用的 path

    '// initial
    bHaveError = Fasle
    outputMessage = ""
    sourceFile = Server.MapPath(xPath & xFile)
    outputFilePath = sourceFile & ".bak"

    Set myStream = Server.CreateObject("ADODB.Stream")
    myStream.Open
    myStream.Type = 1

    Set FSO = Server.CreateObject("Scripting.FileSystemObject")


    '// processing

    '// step 1: check file exist.
    if isFileExist(sourceFile) = 0 then
        outputMessage = "Error:" & sourceFile & " does not exist!"
        bHaveError = True
    end if

    if not bHaveError then
        '// step 2: move data to .bak file.
        call maxCopyFile(sourceFile,outputFilePath)
        call maxDeleteFile(sourceFile)
  
        '// step 3: get file info.
        Set myFile = fso.GetFile(outputFilePath)
        intFilelength = myFile.size
        myStream.LoadFromFile(outputFilePath)

        '// step 4: decode .bak file each byte to source file.
        totalOutputByte = ""
        for i = 1 to intFilelength
            outputByte = ""
            outputByte = chrb(ascb(myStream.Read(1)) xor myPrivateKey)
            'Response.BinaryWrite outputByte
            totalOutputByte=totalOutputByte & outputByte
        next

        outputMessage = saveDataToFile(totalOutputByte, sourceFile)
    end if

    set myStream = nothing
    set myFile = nothing
    Set FSO = Nothing

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
    outputMessage = replace(outputMessage,"<br/>","")
    outputMessage = replace(outputMessage,"\","\\")
    outputMessage = replace(outputMessage,vbCrLf,"\n")
    response.write "alert('" & outputMessage & "');" & vbcrlf
end if
%>
    window.location.href="fileMan.asp?xPath=<%=filemanxpath%>";
</script>
</body>
</html>

<%

'//================================================================================
'//================================================================================

'// purpose: output data to file
'// ex: ret = saveDataToFile(totalOutputByte, outputFilePath)
function saveDataToFile(byval totalOutputByte, byval outputFilePath)
    dim sql
    Dim Blob, BlobData          '// process binary file object.
    dim returnMessage


    Set Blob = Server.CreateObject("TABS.Blob")
    returnMessage = ""

    '// insert binary data to database
    sql = "insert into tabsUploadFile values('" & ezPkStr(sourceFile) & "',0x" & BinaryToHex(totalOutputByte) & "," & session.sessionid & ");"
    'response.write "<br/>sql: " & sql & vbcrlf
    conn.execute sql

    '// read binary data from database
    sql = "select filedata from tabsUploadFile where sessionid=" & session.sessionid
    'response.write "<br/>sql: " & sql & vbcrlf

    set Rs = conn.execute(sql)
    if not rs.eof then
        BlobData = Blob.LoadFromAdoField(rs(0))
        Blob.SaveToFile BlobData, outputFilePath, false
        returnMessage = "<br/>Decode done!" & vbcrlf
        returnMessage = returnMessage & "<br/>File: " & outputFilePath  & vbcrlf

        '// delete binary data from database
        sql = "delete from tabsUploadFile where sessionid=" & session.sessionid
        'response.write "<br/>sql: " & sql & vbcrlf
        conn.execute sql
    else
        returnMessage = "<br/>database recordset not found!"
        'response.write "<br/>sql: " & sql & vbcrlf
    end if

    saveDataToFile = returnMessage
end function

function ezPkStr(byval inputString)
    ezPkStr = replace(inputString,"'","''")
end function

Function BinaryToHex(byval Binary)
  Dim c1, Out, OneByte

  'For each source byte
  For c1 = 1 To LenB(Binary)
    'Get the byte As hex
    OneByte = Hex(AscB(MidB(Binary, c1, 1)))

    'append zero For bytes < 0x10
    If Len(OneByte) = 1 Then OneByte = "0" & OneByte

    'join the byte To OutPut stream
    Out = Out & OneByte
  Next

  'Set OutPut value
  BinaryToHex = Out
End Function


'// purpose: delete a file
'// ex: ret = maxDeleteFile(inputFilepath)
function maxDeleteFile(byval inputFilepath)
    dim deleteFSO
    dim returnValue

    returnValue = 0
    Set deleteFSO = Server.CreateObject("Scripting.FileSystemObject")
    If deleteFSO.FileExists(inputFilepath) Then
        on error resume next
        deleteFSO.DeleteFile(inputFilepath)
    end if

    maxDeleteFile = returnValue
end function



'// purpose: copy source filename to target filename
'// return:
'//     True: success.
'//     False: fail.
'// ex: ret = maxCopyFile(sourceFile,targetFile)
function maxCopyFile(byval sourceFile, byval targetFile)
    dim copyFSO
    dim returnValue
    
    returnValue = Fasle
    if targetFile <> "" and sourceFile <> "" then
        Set copyFSO = Server.CreateObject("Scripting.FileSystemObject")

        If copyFSO.FileExists(sourceFile) Then
            copyFSO.CopyFile sourceFile, targetFile
        end if
        Set copyFSO = Nothing
    end if

    maxCopyFile = returnValue
end function


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
    set FSOCheck = nothing

    isFileExist = returnValue
end function
%>