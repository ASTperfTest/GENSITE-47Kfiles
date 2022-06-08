<%@ CodePage = 65001 %><%
response.charset = "utf-8"

dim fso
dim Folder
dim xPath
dim upPath
dim UploadCount

dim bHaveError
dim strErrorMessage

const defaultUploadCount = 3

'// requset
xPath = request("xPath")
UploadCount = request("uploadCount")

'// initial
Set fso = server.CreateObject("Scripting.FileSystemObject")

upPath = session("uploadPath") & xPath
if right(upPath,1)<>"/" then
    upPath = upPath & "/"
end if
if right(xPath,1)<>"/" then
    xPath = xPath & "/"
end if
if left(xPath,1)="/" then
    xPath = mid(xPath,2)
end if

'// upload counter = 空時, 預設 show 3 筆.
if UploadCount="" or not isnumeric(UploadCount) then
    UploadCount = defaultUploadCount
end if


if "" & session("uploadPath") = "" then
    bHaveError=True
    strErrorMessage = "太久沒動作, 請重新登入: "
    strErrorMessage = strErrorMessage & "<a href='http://" & Request.ServerVariables("SERVER_NAME") & "' target='_top'>" & Request.ServerVariables("SERVER_NAME") & "<a>"
end if

%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=<%=response.charset%>">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<head>
</head>
<body>
<%
if not bHaveError then
%>
<table cellpadding="3" cellspacing="1" border="1" width="98%">
  <tr>
    <td  width="100%" colspan="2"  bgcolor="lightblue" align="center">
        目前路徑: <font color='Red'><%=session("uploadPath")%><%=xpath%></font>
    </td>
  </tr>
  <tr>
    <th bgcolor="lightblue" align="center">上 傳 檔 案</th>
    <th bgcolor="lightblue" align="center">檔 案 清 單</th>
  </tr>
  <tr>
    <td valign="top"> <form action="upload.asp" method="post" enctype="multipart/form-data" ID="Form2">
	<input type="hidden" name="xxPath" value="<%=xpath%>" />
        <table width="100%">
          <tr>
            <td colspan="2" align="left">
              <input type="reset" class="cbutton" value="重設" border="0"  name="reset2" ID="Reset2">
              <input type="submit" class="cbutton" value="送出" border="0" name="send2" id="send2">
              <input type="button" class="cbutton" value="自訂上傳檔案數" name="btnMultiUpload" onclick="javascript:uploadNfiles();">
            </td>
          </tr>

          <% for i = 0 to UploadCount -1 %>
          <tr>
            <td bgcolor="e1e0ca" align="center"><font size='2'>檔案
              <%=i+1%></font></td>
            <td ><input name="photo<%=i%>" type="file" class="cbutton" value=""></td>
          </tr>
          <% next %>
          <tr>
            <td colspan="2" align="left">
              <input type="reset" class="cbutton" value="重設" border="0"  name="reset" ID="Reset1">
              <input type="submit" class="cbutton" value="送出" border="0" name="send" id="send1">
            </td>
          </tr>
        </table>
      </form></td>
    <td valign="top">
        <form name="reg" method="POST" action="delProcess.asp">
        <input type="hidden" name="uploadPath" value="<%=session("uploadPath")%>">
        <input type="hidden" name="xxPath" value="<%=xpath%>">
        <input type="hidden" name="submitTask" value="">
        <input type="button" class="cbutton" name="btnGoParent" value="上一層目錄" onclick="javascrip:goParent();" />
        <input type="button" class="cbutton" value="新增資料夾" name="btnAddFolder" onclick="javascript:addFolder();">
<%
    Set Folder = FSO.GetFolder(server.MapPath(upPath))
    dim totalFileSize
    dim totalFolderSize

    totalFolderSize = 0
    totalFileSize = 0

    if Folder.Files.count > 0 OR Folder.subFolders.count > 0 then
        if Folder.Files.count > 0 then
        %>
        <INPUT type="button" class="cbutton" name="delFile" VALUE="刪檔" onClick="javascript:formDelFileSubmit()" />
        <%
        end if
%>
        <table width="100%">
          <tr bgcolor="lightblue">
            <td align='center'><input type="checkbox" name="checkSelectAll" value="" onclick="javascrip:doSelectAll(document.reg.insertFiles,document.reg.checkSelectAll.checked);"/></td>
            <td>檔 名</td>
            <td align="center">大 小</td>
            <td align="center">修改日期</td>
            <td align="center"></td>
          </tr>
          <%

            '// list folder
            For Each FolderDir In Folder.SubFolders
                FolderDate = RDT(FolderDir.DateLastModified)
                FolderSize = ReplaceSize(FolderDir.Size)
                totalFolderSize = totalFolderSize + FolderDir.Size
            %>
          <tr bgcolor="yellow">
          <td>&nbsp;</td>
          <td><a href="fileMan.asp?xpath=<%=server.urlencode(xpath & FolderDir.Name)%>"><%=getFolderImageIcon(FolderDir.Name) & "&nbsp;" %>
          <%=FolderDir.Name%></a></td>
          <td align='right'><font size='2'><%=FolderSize%></font></td>
          <td align='right'><font size='2'><%=FolderDate%></font></td>
          <td align='right'>
                <input type='button' class="cbutton" name='renfile_<%=i%>' value='修改' onclick="javscript:renamedir('<%=FolderDir.name%>');" />
          </td>
          </tr>
          <%
            Next


            '// list files
            i=0
            for each FolderFiles in Folder.Files
                i=i+1
                FileDate = RDT(FolderFiles.DateLastModified)
                FileSize = ReplaceSize(FolderFiles.Size)
                totalFileSize = totalFileSize + FolderFiles.Size
%>
          <tr <%
            if i mod 2 = 0 then
                response.write " bgcolor='#e0e0f0' "
            end if %>>
            <td width='10' nowrap='true' align='center'><input type="checkbox" name="insertFiles" value="<%=FolderFiles.name%>" /></td>
            <td><A target="_blank" href="/xfup/maxGet.asp?xfile=<%=session("uploadPath")%><%=xpath%><%=FolderFiles.name%>"><%=getFileImageIcon(FolderFiles.name) & "&nbsp;" %>
                <%=FolderFiles.name%>
              </A></td>
              <td align='right'><font size='2'><%=FileSize%></font></td>
              <td align='right'><font size='2'><%=FileDate%></font></td>
              <td align='right'>
                <input type='button' class="cbutton" name='renfile_<%=i%>' value='修改' onclick="javscript:renamefile('<%=FolderFiles.name%>');" />
                <input type='button' class="cbutton" name='delfile_<%=i%>' value='刪除' onclick="javscript:deleteOneFile('<%=FolderFiles.name%>');" />
                <input type='button' class="cbutton" name='decodefile_<%=i%>' value='解碼' onclick="javscript:doFileDecode('<%=FolderFiles.name%>');" />
              </td>


          </tr>
          <%
                next
            %>
        </table>
        <%
        if Folder.Files.count > 0 then
        %>
        <INPUT type="button" class="cbutton" name="delFile2" VALUE="刪檔" onClick="javascript:formDelFileSubmit()">
        <br/>
        <%
        end if
        %>

      <div align='right'><font color='#808080'>
      <%
   Response.Write "<br/>共有 <em><b>" & Folder.subFolders.count & "</b></em> 個資料夾和 <em><b>" & Folder.Files.count & "</b></em> 個檔案" & VbCrlf
   Response.Write ", 整個資料夾共佔用 <em><b>" & ReplaceSize(totalFileSize + totalFolderSize) & "</b></em> " & VbCrlf
   Response.Write "<br/>其中資料夾共佔用 <em><b>" & ReplaceSize(totalFolderSize) & "</b></em> " & VbCrlf
   Response.Write ", 檔案共佔用 <em><b>" & ReplaceSize(totalFileSize) & "</b></em> " & VbCrlf
   %>
        </font></div>

          <%  else %>
            <INPUT type="button" class="cbutton" name="delDir" VALUE="刪除此目錄" onClick="javascript:formDelDirSubmit()">
          <%  end if %>
      </form>
      </td>
  </tr>
</table>


<form action="upload.asp" method="post" enctype="multipart/form-data" ID="FormAddFolder" name="FormAddFolder">
	<input type="hidden" name="xxPath" value="<%=xpath%>" />
    <input type="hidden" name="newFolder" size=20 value="" />
    <input type="hidden" name="send" value="新增" />
</form>

<form name="FormRenFile" method="post" action="maxRenFile.asp">
	<input type="hidden" name="filemanxpath" value="<%=xpath%>" />
	<input type="hidden" name="xPath" value="<%=session("uploadPath")%><%=xpath%>" />
    <input type="hidden" name="xFile" value="" />
    <input type="hidden" name="NewFileName" value="" />
    <input type="hidden" name="renameMode" value="" />
</form>

<form name="deleteOne" method="POST" action="delProcess.asp">
    <input type="hidden" name="uploadPath" value="<%=session("uploadPath")%>">
    <input type="hidden" name="xxPath" value="<%=xpath%>">
    <input type="hidden" name="submitTask" value="">
    <input type="hidden" name="insertFiles" value="">
</form>

<form name="fileDecode" method="POST" action="maxFileDecode.asp">
	<input type="hidden" name="filemanxpath" value="<%=xpath%>" />
	<input type="hidden" name="xPath" value="<%=session("uploadPath")%><%=xpath%>" />
    <input type="hidden" name="xFile" value="" />
</form>

</body>
<script language=javascript>
function formDelFileSubmit(){
    if(confirm("注意！\n\n您確定刪除檔案嗎？")){
        document.reg.submitTask.value = "delFiles";
        document.reg.submit();
    }
}

function formDelDirSubmit(){
    if(confirm("注意！\n\n您確定刪除目錄嗎？")){
        document.reg.submitTask.value = "delDir";
        document.reg.submit();
    }
}
</script>

<script language="javascript">
<!--
// purpose: select all
// ex: doSelectAll(fieldObject,true)
function doSelectAll(field, flag) {
    if (typeof(field.length) != 'undefined')
    {
        for (i = 0; i < field.length ; i++)
        {
            field[i].checked = flag;
        }
    }
    else
    {
        field.checked = flag;
    }
}


// purpose: 自訂上傳檔案數
function uploadNfiles(){
    var xpath='<%=server.urlencode(xpath)%>';
    var uploadCount = prompt("請輸入要上傳的檔案筆數:",'');
    if(uploadCount!=null && uploadCount!=''){
    window.location.href="fileMan.asp?xpath="+xpath+"&uploadCount="+uploadCount+"";
    }
}


// purpose: 新增資料夾
function addFolder(){
    var xpath='<%=server.urlencode(xpath)%>';
    var newFolder = prompt("請輸入資料夾名稱:",'');
    if(newFolder!=null && newFolder!=''){
        document.FormAddFolder.newFolder.value=newFolder;
        document.FormAddFolder.submit();
    }
}


// purpose: rename file / folder
function renamefile(oldfilename){
    var newfilename=prompt("請輸入文件 ["+oldfilename+"] 的新名稱:",oldfilename);
    if(newfilename!=null && oldfilename!='' && newfilename!=oldfilename){
        document.FormRenFile.renameMode.value="file";
        document.FormRenFile.xFile.value=oldfilename;
        document.FormRenFile.NewFileName.value=newfilename;
        document.FormRenFile.submit();
    }
}

function renamedir(oldfilename){
    var newfilename=prompt("請輸入資料夾 ["+oldfilename+"] 的新名稱:",oldfilename);
    if(newfilename!=null && oldfilename!='' && newfilename!=oldfilename){
        document.FormRenFile.renameMode.value="dir";
        document.FormRenFile.xFile.value=oldfilename;
        document.FormRenFile.NewFileName.value=newfilename;
        document.FormRenFile.submit();
    }
}


// purpose: delete file.
function deleteOneFile(oldfilename){
    if(oldfilename!=null && oldfilename!=''){
        if(confirm("您確定要刪除 " + oldfilename + " 嗎?")){
            document.deleteOne.insertFiles.value=oldfilename;
            document.deleteOne.submitTask.value='delFiles';
            document.deleteOne.submit();
        }
    }
}


// purpose: Decode file
function doFileDecode(oldfilename){
    if(oldfilename!=null && oldfilename!=''){
        document.fileDecode.xFile.value=oldfilename;
        document.fileDecode.submit();
    }
}



function goParent(){
<%
dim myParentPath
myParentPath = xpath
if right(myParentPath,1) = "/" then
    myParentPath =left(myParentPath,len(myParentPath)-1)
end if
myParentPath = trimLastDelimit(myParentPath,"/")
if right(myParentPath,1) <> "/" then
    myParentPath = myParentPath & "/"
end if
if myParentPath = xpath then
    myParentPath = ""
end if
%>
    var xpath='<%=server.urlencode(myParentPath)%>';
    window.location.href="fileMan.asp?xpath="+xpath;
}
-->
</script>
<%
else
    response.write strErrorMessage
end if
%>
</body>
</html>
<%

set fso = nothing
set conn = nothing



'// 用於顯示 易於讀取的 檔案大小.
Function ReplaceSize(FileSize)

    If FileSize < 1024 then
        ReplaceSize = left(FileSize,InStr(FileSize,".") + 2) & " Bytes"
        Exit Function
    End If

    FileSize = FileSize / 1024

    If FileSize < 1024 then
        ReplaceSize = left(FileSize,InStr(FileSize,".") + 2) & " KB"
        Exit Function
    End If

    FileSize = FileSize / 1024

    If FileSize < 1024 then
        ReplaceSize = left(FileSize,InStr(FileSize,".") + 2) & " MB"
        Exit Function
    End If

    FileSize = FileSize / 1024
    ReplaceSize = left(FileSize,InStr(FileSize,".") + 2) & " GB"
    'TB不計(現在一般硬碟未有如此大的空間)
End Function


'// 用於顯示 易於讀取的 檔案日期/時間.
Function RDT(DateTime)
    RDT = "yyyy/mm/dd AMPM hh:nn:ss"
    DateTime = CDate(DateTime)
    Y = Right("0000" & DatePart("yyyy" ,DateTime),4)	'取出年
    M = Right("0" & DatePart("m" ,DateTime),2)		'取出月
    D = Right("0" & DatePart("d" ,DateTime),2)		'取出日
    H = DatePart("h" ,DateTime)		'取出時

    if instr(RDT,"AMPM") <> 0 then
        AMPM = "AM"

        if H > 12 then
            H = H - 12
            AMPM = "PM"
        End if

    end if

    H = Right("0" & H,2)
    N = Right("0" & DatePart("n" ,DateTime),2)		'取出分
    S = Right("0" & DatePart("s" ,DateTime),2)		'取出秒

    RDT = Replace(RDT , "yyyy" , Y)
    RDT = Replace(RDT , "mm" , M)
    RDT = Replace(RDT , "dd" , D)
    RDT = Replace(RDT , "AMPM" , AMPM)
    RDT = Replace(RDT , "hh" , H)
    RDT = Replace(RDT , "nn" , N)
    RDT = Replace(RDT , "ss" , S)
End Function


'// purpose: 去掉字串裡, 某個符號右邊的字元.
'// input:
'//     inputString, 字串
'//     myDelimit, 某個符號
'// ps: "特定符號" 可以是多個字元, 例如 "&amp;",
'// ex: ret = trimLastDelimit(inputString,Delimit)
function trimLastDelimit(byval inputString, byval myDelimit)
    dim returnValue
    dim tempIndex
    dim tempString

    returnValue = inputString
    if inputString <> "" and myDelimit <> "" then
        if instr(inputString,myDelimit) > 0 then
            tempString=StrReverse(inputString)
            myDelimit = StrReverse(myDelimit)
            tempIndex = len(tempString) - instr(tempString,myDelimit)
            returnValue = left(inputString,tempIndex - len(myDelimit) + 1)
        end if
    end if
    trimLastDelimit = returnValue
end function


'// purpose: get icon
'// return: image filepath.
'// ex: ret = getFolderImageIcon(filename)
function getFolderImageIcon(byval filename)
    dim returnValue
    dim imageFolder
    dim FileIcon
    dim HelpStr

    imageFolder = "/xfup/icons/"
    'FileIcon = "dir.gif"
    FileIcon = "folder.gif"
    HelpStr = "資料夾"

    if FileIcon <> "" then
        returnValue = imageFolder & FileIcon
    end if

    returnValue = "<img src='" & returnValue & "' alt='" & filename & " (" & helpStr & ")' border='0' align='middle' />"

    getFolderImageIcon = returnValue
end function


'// purpose: get icon
'// return: image filepath.
'// ex: ret = getFileImageIcon(filename)
function getFileImageIcon(byval filename)
    dim returnValue
    dim imageFolder
    dim FileIcon
    dim HelpStr
    dim caseCondition

    imageFolder = "/xfup/icons/"
    FileIcon = ""
    HelpStr = ""
    caseCondition = LCase(Right(FileName,Len(FileName)-Instrrev(FileName,".")))

    Select Case caseCondition
    Case "ace","arj","bz2","cab","iso","jar","lzh","rar","tar","uue"
        FileIcon = "rar.gif"
        HelpStr = "WinRAR 壓縮檔"

    Case "asp","aspx", "inc"
        FileIcon = "asp.gif"
        HelpStr = "Active Server Pages 文件"

    Case "avi", "mkv", "mov", "wmv"
        FileIcon = "mov.gif"
        HelpStr = "視訊短片"

    Case "bak"
        FileIcon = "unknown.gif"
        HelpStr = "備份檔"

    Case "bat"
        FileIcon = "exe.gif"
        HelpStr = "Batch 批次檔"

    Case "bmp"
        FileIcon = "bmp.gif"
        HelpStr = "點陣圖影像"

    Case "c"
        FileIcon = "c.gif"
        HelpStr = "C File"

    Case "chm"
        FileIcon = "chm.gif"
        HelpStr = "已編譯的 HTML Help 檔案"

    Case "com" , "exe"
        FileIcon = "exe.gif"
        HelpStr = "執行檔"

    Case "config",  "asa" , "asax"
        FileIcon = "text.gif"
        HelpStr = "Config 設定文件"

    Case "cgi","pl"
        FileIcon = "unknown.gif"
        HelpStr = "Perl 檔案"

    Case "css"
        FileIcon = "css.gif"
        HelpStr = "Cascading Style Sheet 文件"

    Case "db"
        FileIcon = "dll.gif"
        HelpStr = "資料庫檔案"

    Case "dll"
        FileIcon = "dll.gif"
        HelpStr = "應用程式擴充"

    Case "doc"
        FileIcon = "word.gif"
        HelpStr = "Microsoft Word 文件"

    Case "gif"
        FileIcon = "gif.gif"
        HelpStr = LCase(Right(FileName,Len(FileName) - Instrrev(FileName,"."))) & " 影像"

    Case "htm","html"
        FileIcon = "html.gif"
        HelpStr = "HTML Document"

    Case "ini"
        FileIcon = "ini.gif"
        HelpStr = "組態設定值"

    Case "jpg","jpeg"
        FileIcon = "jpg.gif"
        HelpStr = LCase(Right(FileName,Len(FileName) - Instrrev(FileName,"."))) & " 影像"

    Case "js"
        FileIcon = "js.gif"
        HelpStr = "JavaScript 檔案"

    Case "mdb"
        FileIcon = "access.gif"
        HelpStr = "Access 資料庫"

    Case "mid","midi"
        FileIcon = "midi.gif"
        HelpStr = "MIDI Sequence"

    Case "mp3"
        FileIcon = "mp3.gif"
        HelpStr = "MP3 格式聲音"

    Case "mpeg"
        FileIcon = "mpeg.gif"
        HelpStr = "電影檔 (MPEG)"

    Case "pdf"
        FileIcon = "pdf.gif"
        HelpStr = "Adobe Acrobat Document"

    Case "php"
        FileIcon = "text.gif"
        HelpStr = "PHP 檔案"

    Case "png"
        FileIcon = "png.gif"
        HelpStr = LCase(Right(FileName,Len(FileName) - Instrrev(FileName,"."))) & " 影像"

    Case "ppt"
        FileIcon = "ppt.gif"
        HelpStr = "PowerPoint 投影片"

    Case "ra", "rm", "rmvb"
        FileIcon = "realplayer.gif"
        HelpStr = "RealPlayer 多媒體檔案"

    Case "svn", "svn-base"
        FileIcon = "svn.gif"
        HelpStr = "SVN 檔案"

    Case "swf", "fls"
        FileIcon = "swf.gif"
        HelpStr = "Flash 檔案"

    Case "txt", "sql", "log", "text", "csv"
        FileIcon = "text.gif"
        HelpStr = "純文字文件"

    Case "vbs"
        FileIcon = "vbs.gif"
        HelpStr = "VBScript 檔案"

    Case "wav"
        FileIcon = "wav.gif"
        HelpStr = "聲音"

    Case "xml"
        FileIcon = "xml.gif"
        HelpStr = "Extensible Markup Language"

    Case "xls" 
        FileIcon = "excel.gif"
        HelpStr = "Extensible Stylesheet Language"

    Case "xsl"
        FileIcon = "xsl.gif"
        HelpStr = "Extensible Stylesheet Language"

    Case "zip", "gzip"
        FileIcon = "zip.gif"
        HelpStr = "Zip 壓縮檔"

    Case Else
        if instr(caseCondition,".") > 0 then
            FileIcon = "unknown.gif"
        else
            FileIcon = "unknown2.gif"
        end if
        HelpStr = "檔案" & caseCondition
    End Select

    if FileIcon <> "" then
        returnValue = imageFolder & FileIcon
    end if

    returnValue = "<img src='" & returnValue & "' alt='" & server.mappath(session("uploadPath") & xpath) & "\" & filename & " (" & helpStr & ")' border='0' width='18' align='middle' />"

    getFileImageIcon = returnValue
end function
%>
