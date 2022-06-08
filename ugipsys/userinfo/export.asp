<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="使用者登入系統記錄留存備查 "
HTProgFunc="使用者登入訊息"
HTProgCode="BM010"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->

<!--#include FILE = "../inc/selectTree.inc" -->
<!--#include FILE = "../HTSDSystem/HT2CodeGen/htUIGen.inc" -->

<%
queryuser=request("select0")

if queryuser<>"全部使用者" then
if request("htx_CuDTx23F84S")<>"" and request("htx_CuDTx23F84E")<>"" then
sql2="exec sp_userinfo N'" & queryuser & "',N'" & request("htx_CuDTx23F84S") & "',N'" & dateadd("d",1,request("htx_CuDTx23F84E")) & "'"
else
sql2="exec sp_userinfo N'" & queryuser & "','',''"
end if
else
if request("htx_CuDTx23F84S")<>"" and request("htx_CuDTx23F84E")<>"" then
sql2="exec sp_userinfo '',N'" & request("htx_CuDTx23F84S") & "',N'" & dateadd("d",1,request("htx_CuDTx23F84E")) & "'"
else
sql2="exec sp_userinfo '','',''"
end if
end if
        
               Set rsData = conn.execute(sql2)
               If Not rsData.EOF Then
                  Set fs= Server.CreateObject("Scripting.FileSystemObject")
                  FileName = UploadTempPath & Replace(request("htx_CuDTx23F84S"),"/","_") & "~" & Replace(request("htx_CuDTx23F84E"),"/","_") & ".csv"
                  ExportFile = Server.MapPath(FileName)
                  Set txtf = fs.CreateTextFile(ExportFile)
   
                  DataTemp = "姓名,帳號,登入IP,時間,狀態,主題單元,異動資料 "
                  txtf.WriteLine DataTemp
             
                  While Not rsData.EOF
                     DataTemp = Trim(rsData("UserName")) & "," & Trim(rsData("userid")) &"," & Trim(rsData("loginip")) & "," & Trim(rsData("logintime")) & "," & Trim(rsData("xtarget")) & "," & Trim(rsData("ctunit")) & "," & Trim(rsData("objtitle")) 

                     txtf.WriteLine DataTemp
             
                     rsData.MoveNext
                  Wend
                  Response.Redirect FileName
               Else
                  Response.Write "<font color='#993300'>本段日期無資料可供匯出，請重新輸入！</font><br>"
                  flag = False
               End If
         
         If flag = False Then
%>
<form action="index.asp?page=1" method="post">
   <input type="hidden" name="send" value="">
   <input type="submit" value="資料匯出尚未完成，回上一頁！">
</form>
<%
         End If
     
%>