<% Response.Expires = 0
HTProgCap="���ʾǭ��޲z"
HTProgFunc="�M��"
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
   
                  DataTemp = "�s��,�����Ҹ�,�m�W,�X�ͤ�,�ʧO,eMail,�s���q��,�p���a�},��¾����,¾��,�̰��Ǿ�,�����ߺD,�A�Ⱦ��c,���q�Τ@�s��,�줽�a�},���q�q��,�ǯu,�����ӷ�,��ʹq��,�ѻP�ҵ{"
                  txtf.WriteLine DataTemp
                  counter1=1
                  While Not rsData.EOF
                         sql="select c.actName from paEnroll a,paSession b,ppAct c where a.psnID='" &  Trim(rsData("psnID")) & "' and a.paSID=b.paSID and b.actID=c.actID"
                         set rs=conn.Execute(sql)
                           
                     DataTemp = counter1 & "," & Trim(rsData("psnID")) & "," & Trim(rsData("pName")) &"," & Trim(rsData("birthDay")) & "," & Trim(rsData("sex")) & "," & Trim(rsData("eMail")) & "," & Trim(rsData("tel")) & "," & Trim(rsData("emergContact")) & "," & rsData("deptName") & "," & rsData("jobName")& "," & rsData("topEdu")& "," & rsData("meatKind")& "," & Trim(rsData("corpName")) & "," & rsData("corpID") & "," & rsData("corpAddr") & "," & rsData("corpTel")& "," & rsData("corpFax")& "," & rsData("datafrom")& "," & rsData("mobile")& "," 
                     do while not rs.eof
                          DataTemp=DataTemp & trim(rs("actName"))& "�B"
                     rs.MoveNext
                     loop
                     txtf.WriteLine DataTemp
                     counter1=counter1+1
                     rsData.MoveNext
                  Wend
                  Response.Redirect FileName
               Else
                  Response.Write "<font color='#993300'>�L��ƥi�ѶץX�I</font><br>"
                  flag = False
               End If
          
%>


