<% @ CodePage = 65001 %>
<%
'// purpose: 為了 emily 去 download Moj 裡 any type of files from server.
'// date: 2006/7/13

Response.Buffer = True
Response.Clear

HTProgCode = "GC1AP5"
CatalogueFrame = "220,*"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
dim xFile
dim downFile
dim FSO

xFile = request("xFile")
DownFile = Server.MapPath(xFile)

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

FileFind = FSO.FileExists(DownFile)

If Not FileFind Then
    Response.Write("<h1>Error:</h1>" & DownFile & " does not exist<p>")
    Response.End
End If

If Request("Act") = "save" Then
        Set txtf = FSO.OpenTextFile(DownFile,2,True)
        txtf.Write Request("file")
        Set txtf = Nothing
        Response.write  "檔案" & downfile & "編輯或新增完成！"
        Response.end
End If

%>
<HTML>
<HEAD>
<title>編輯<%=DownFile%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<Script Language="JavaScript">
<!--
function filebody(){
<%
   If FileFind Then
     Set txtf = FSO.OpenTextFile(DownFile,1)
     AtFirst=True
     Do Until txtf.AtEndOfStream
       txtLine = txtf.ReadLine
       txtLine = Replace(txtLine,"\","\\")
       txtLine = Replace(txtLine,"""","\""")
       txtLine = Replace(txtLine,"<","<""+""")
       If AtFirst Then
         if asc(left(txtline,1)) = 0 then
            txtline = right(txtline,len(txtline)-1)
            if left(txtline,2) = "?%" then
                txtline = "<""+""" & right(txtline,len(txtline)-1)
            end if
         end if
         response.write " // the first char is:" & left(txtline,1)
         response.write " // the first char is:" & asc(left(txtline,1))
         response.write " // the first 2 char is:" & left(txtline,2)
	    Response.Write """;" & VbCrlf
        
         Response.Write "  document.frm.file.value=""" & txtLine
         AtFirst = False
       Else
         Response.Write "  document.frm.file.value+=""" & txtLine
       End If
       If txtf.AtEndOfStream Then
	 Response.Write """;" & VbCrlf
       Else
	 Response.Write "\n"";" & VbCrlf
       End If
       'If Not txtf.AtEndOfStream Then Response.Write VbCrLf
     Loop
     Set txtf = Nothing
   End If
%>
}
-->
</Script>
</HEAD>
<BODY onload="filebody();">

<div align="center">
  <center>
  <form method="POST" action="maxEditFile.asp" name="frm">

  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%">
    <tr>
      <td width="100%"><%If FileFind Then%>編輯於<%Else%>建立於<%End If%>　
      <input type="text" name="xFile" value="<%=xFile%>" class="tx2" style="width:40%" size="20" maxlength="250" readonly="true"><BR><center>
       <textarea rows="2" name="file" cols="20" style="width:100%; height:500" class="tx2"></textarea><br>
        <input type="submit" value="存檔" name="B1" class="button2">
        <input type="reset" value="重寫" name="B2" class="button2">
       </center></td>
    </tr>
  </table>
  <input type="hidden" name="Act" value="save">
  </form>
  </center>
</div>
</BODY>
</HTML>