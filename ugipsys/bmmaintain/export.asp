<%@ CodePage = 65001 %>
<!--#Include file = "../inc/server.inc" -->
<%
Function CheckDateOK( fieldY, fieldM, fieldD )
   Session("DateOK") = True

   If IsNumeric(fieldY)=True AND IsNumeric(fieldM)=True AND IsNumeric(fieldD)=True Then
      If CInt(fieldM)<1 OR CInt(fieldM)>12 Then
         Session("DateOK") = False
      Else
         If CInt(fieldM)=1 OR CInt(fieldM)=3 OR CInt(fieldM)=5 OR CInt(fieldM)=7 OR CInt(fieldM)=8 OR CInt(fieldM)=10 OR CInt(fieldM)=12 Then
            IF CInt(fieldD)<1 OR CInt(fieldD)>31 Then
               Session("DateOK") = False
            End If
         ElseIf CInt(fieldM)=4 OR CInt(fieldM)=6 OR CInt(fieldM)=9 OR CInt(fieldM)=11 Then
            IF CInt(fieldD)<1 OR CInt(fieldD)>30 Then
               Session("DateOK") = False
            End If
         ElseIf CInt(fieldM) = 2 Then
            If (CInt(fieldY)+1911) Mod 4 = 0 Then
               If CInt(fieldD)<1 OR CInt(fieldD)>29 Then
                  Session("DateOK") = False
               End If
            Else
               If CInt(fieldD)<1 OR CInt(fieldD)>28 Then
                  Session("DateOK") = False
               End If
            End If
         End If
      End If
   Else
      Session("DateOK") = False
   End IF
End Function

         flag = True
         CheckDateOK Request("textfield2"),Request("textfield22"),Request("textfield23")
         If Session("DateOK") = True Then
            Temp = 1911 + CInt(Request("textfield2"))
            StartDay = DateSerial(Temp,Request("textfield22"),Request("textfield23"))
         Else
            Response.Write "<font color='#993300'>開始日期應為有效數字，請重新輸入！</font><br>"
            flag = False
         End If

         CheckDateOK Request("textfield24"),Request("textfield25"),Request("textfield26")
         If Session("DateOK") = True Then
            Temp = 1911 + CInt(Request("textfield24"))
            EndDay = DateSerial(Temp,Request("textfield25"),Request("textfield26"))
         Else
            Response.Write "<font color='#993300'>結束日期應為有效數字，請重新輸入！</font><br>"
            flag = False
         End If

         If flag = True Then
            If StartDay > EndDay Then
               Response.Write "<font color='#993300'>日期先後錯誤，請重新輸入！</font><br>"
               flag = False
            Else
               SQLData = "Select * From Prosecute Where Date>='" & StartDay & " 0:0:0' AND Date<='" & EndDay & " 23:59:59' Order By Date"
               Set rsData = conn.execute(SQLData)
               If Not rsData.EOF Then
                  Set fs= Server.CreateObject("Scripting.FileSystemObject")
                  FileName = UploadTempPath & Replace(StartDay,"/","_") & "~" & Replace(EndDay,"/","_") & ".csv"
                  ExportFile = Server.MapPath(FileName)
                  Set txtf = fs.CreateTextFile(ExportFile)
   
                  DataTemp = "編號,類別,文號,來信日期,姓名,聯絡電話,電子郵件信箱,聯絡地址,意見內容,回應單位,回應內容,回應日期,回應期限"
                  txtf.WriteLine DataTemp
                  counter1=1
                  While Not rsData.EOF
                     DataTemp = counter1 & "," & Trim(rsData("classname")) & "," & Trim(rsData("docid")) &"," & Trim(rsData("Date")) & "," & Trim(rsData("Name")) & "," & Trim(rsData("Tele")) & "," & Trim(rsData("Email")) & "," & Trim(rsData("Addr")) & "," & Trim(rsData("Context")) & "," & rsData("unit") & "," & Trim(rsData("Reply")) & "," & rsData("ReplyDate") & "," & rsData("dateline")

                     txtf.WriteLine DataTemp
                     counter1=counter1+1
                     rsData.MoveNext
                  Wend
                  Response.Redirect FileName
               Else
                  Response.Write "<font color='#993300'>本段日期無資料可供匯出，請重新輸入！</font><br>"
                  flag = False
               End If
            End If
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