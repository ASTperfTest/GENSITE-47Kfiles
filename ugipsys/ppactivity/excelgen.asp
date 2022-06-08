<% Response.Expires = 0
HTProgCap="活動學員管理"
HTProgFunc="清單"
HTProgCode="PA010"
HTProgPrefix="ppPsnInfo" %>
<!--#INCLUDE FILE="ppPsnInfoListParam.inc" -->
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 
%>
<%
               SQLData = "Select * From paPsnInfo order by cdate"
               Set rsData = conn.execute(SQLData)
               If Not rsData.EOF Then
                  Set fs= Server.CreateObject("Scripting.FileSystemObject")
                  FileName = UploadTempPath & Replace(StartDay,"/","_") & "~" & Replace(EndDay,"/","_") & ".csv"
                  ExportFile = Server.MapPath(FileName)
                  Set txtf = fs.CreateTextFile(ExportFile)
   
                  DataTemp = "編號,身份證號,姓名,出生日,性別,eMail,連絡電話,聯絡地址,任職部門,職稱,最高學歷,飲食習慣,服務機構,公司統一編號,辦公地址,公司電話,傳真,消息來源,行動電話,參與課程"
                  txtf.WriteLine DataTemp
                  counter1=1
                  While Not rsData.EOF
                         sql="select c.actName from paEnroll a,paSession b,ppAct c where a.psnID='" &  Trim(rsData("psnID")) & "' and a.paSID=b.paSID and b.actID=c.actID"
                         set rs=conn.Execute(sql)
                           
                     DataTemp = counter1 & "," & Trim(rsData("psnID")) & "," & Trim(rsData("pName")) &"," & Trim(rsData("birthDay")) & "," & Trim(rsData("sex")) & "," & Trim(rsData("eMail")) & "," & Trim(rsData("tel")) & "," & Trim(rsData("emergContact")) & "," & rsData("deptName") & "," & rsData("jobName")& "," & rsData("topEdu")& "," & rsData("meatKind")& "," & Trim(rsData("corpName")) & "," & rsData("corpID") & "," & rsData("corpAddr") & "," & rsData("corpTel")& "," & rsData("corpFax")& "," & rsData("datafrom")& "," & rsData("mobile")& "," 
                     do while not rs.eof
                          DataTemp=DataTemp & trim(rs("actName"))& "、"
                     rs.MoveNext
                     loop
                     txtf.WriteLine DataTemp
                     counter1=counter1+1
                     rsData.MoveNext
                  Wend
                  Response.Redirect FileName
               Else
                  Response.Write "<font color='#993300'>無資料可供匯出！</font><br>"
                  flag = False
               End If
          
%>


