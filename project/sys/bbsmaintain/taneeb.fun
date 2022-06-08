
<%
'====================================================================================
' 顯示民國日期  YY/MM/DD
Function ShowDate(Date)

    ShowDate = ShowDate & Year(Date)-1911
    If Month(Date) < 10 Then
        ShowDate = ShowDate & "/0" & Month(Date) 
    Else
        ShowDate = ShowDate & "/" & Month(Date) 
    End If
                
    If Day(Date) < 10 Then 
        ShowDate = ShowDate & "/0" & Day(Date)
    Else
               ShowDate = ShowDate & "/" & Day(Date)
    End If

end function

'====================================================================================
' 轉換 http://  
'      email:
Function TransHttpEmail(Data)

        Dim Point
        Dim Point2
        Dim HTTP
        Point = 1
        Point = InStr(Point, Data, " http://")
        
        While Point > 0
                
                Point2 =InStr(Point+1, Data, " ")
                If Point2 > 1 Then
                        HTTP = Mid( Data, Point+1, Point2-Point-1)
                        'Response.Write HTTP
                        Data = Replace(Data , HTTP, "<a href=" & HTTP & ">" & HTTP & "</a>")                
                        Point = InStr(Point2+1, Data, " http://")                        
                Else
                        Point = 0
                End If
                
        WEnd
        
        Dim EMial
        Point = 1
        Point = InStr(Point, Data, " email:")

        While Point > 0
                T = T + 1
                Point2 =InStr(Point+1, Data, " ")
                If Point2 > 1 Then
                        EMail = Mid( Data, Point+7, Point2-Point-7)
                        Data = Replace(Data , "email:" & EMail, "<a href=mailto:" & EMail & ">" & EMail & "</a>")
                        Point = InStr(Point2+20, Data, " email:")
                Else
                        Point = 0
                End If
                
        WEnd        
 
        TransHttpEmail = Data
End Function

'====================================================================================

function ShowArticle
   
   Dim No
   Dim Date 

   Set rs    = conn.Execute("Select NickName, Article.Date, Article.Id, Article.MId, Article.SId, Article.Bid, Article.Title,article.cnt,isnull(article.isethics,0) isethics  from Article Where  bId = '" & Request("BId") & "' order by MId desc, SId")
   Set rs2    = conn.Execute("Select count(*)  from Article Where  bId = '" & Request("BId") & "' ")
   
   TotalPage = Int(rs2(0)/10)
   If rs2(0) Mod 10 <> 0 Then
        TotalPage = TotalPage + 1
   End if       

   Page= Request("Page")
   If Page = "" Then
        Page = 1        
   Elseif Int(Page) <= 0 Then
        Page = 1    
   Elseif Int(Page)>Int(TotalPage) Then
        Page = TotalPage        
   end if

   No = 1                       
   while Not rs.Eof
        If No >= (Page-1)*10+1 and No <= Page*10 Then
           If No Mod 2 = 1 Then
                Response.Write "<tr class='tr-white'>"
           else
                Response.Write "<tr class='tr-grey'>"
           end if     
                Response.Write "<td>"
                Response.Write "<div align='center'>" & No & "</div>"
                Response.Write "</td>"
                Response.Write "<td width='60%'><a href=Article.asp?MId=" & rs("MID") & "&SId=" & rs("SID") & "&bId=" & Request("BId") & "&Time=" & Timer & ">" & rs("Title") & "</a>"
                if  rs("isethics")="1"  then
                     response.write  "<img src=images/ethics_logo.gif  border=0>"
                end  if   
                response.write  "</td>"
                'Response.Write "<td nowrap><a href=mailto:" & rs("Email") & ">" & rs("Name") & "</a></td>"
                Response.Write "<td nowrap>" & rs("NickName") & "</td>"
                Date = rs("Date")
                Response.Write "<td >" & DateValue(Date) & "　 " & TimeValue(Date) & "</td>"
 		Response.Write "<td nowrap>" & rs("cnt") & "</td>"                
		Response.Write "</tr>"
        End if
        No = No+1
        rs.Movenext   
   Wend
              
end function
'====================================================================================

function ShowArticlesysop
   
   Dim No
   Dim Date 

   Set rs    = GetSQLServerStaticRecordset( conn, "Select NickName, Article.Date, Article.Id, Article.MId, Article.SId, Article.Bid, Article.Title,article.cnt,article.ip,isnull(article.isethics,0) isethics from Article Where  bId = '" & Request("BId") & "' order by MId desc, SId")
   
   TotalPage = Int(rs.Recordcount/10)
   If rs.recordcount Mod 10 <> 0 Then
        TotalPage = TotalPage + 1
   End if       

   Page= Request("Page")
   If Page = "" Then
        Page = 1        
   Elseif Int(Page) <= 0 Then
        Page = 1    
   Elseif Int(Page)>Int(TotalPage) Then
        Page = TotalPage        
   end if

   No = 1                       
   while Not rs.Eof
        If No >= (Page-1)*10+1 and No <= Page*10 Then
                Response.Write "<tr bgcolor='#99CCFF'>"
                Response.Write "<td>"
                Response.Write "<div align='center'>" & No & "</div>"
                Response.Write "</td>"
                Response.Write "<td width='60%'><a href=sysop_Article.asp?MId=" & rs("MID") & "&SId=" & rs("SID") & "&bId=" & Request("BId") & "&Time=" & Timer & ">" & rs("Title") & "</a>"
                 if  rs("isethics")="1"  then
                     response.write  "<img src=images/ethics_logo.gif  border=0>"
                end  if   
                response.write  "</td>"
                'Response.Write "<td nowrap><a href=mailto:" & rs("Email") & ">" & rs("Name") & "</a></td>"
                Response.Write "<td nowrap>" & rs("NickName") & "</td>"
                Date = rs("Date")
                Response.Write "<td ><font size=-1>" & DateValue(Date) & "　 " & TimeValue(Date) & "</font></td>"
 		Response.Write "<td nowrap>" & rs("cnt") & "</td>"   
                Response.Write "<td nowrap>" & rs("ip") & "</td>"               
		Response.Write "</tr>"
        End if
        No = No+1
        rs.Movenext   
   Wend
              
end function
'==========================================================================


Function ShowNewJoe

    Set rs   = GetSQLServerStaticRecordset( conn, "select * from NewJob where Date >= '" & DateValue(Now) & "' Order By PublishDate DESC")
    
        TotalPage = Int(rs.recordcount/5)
        If rs.recordcount Mod 5 <> 0 Then
                TotalPage = TotalPage + 1
        End if
        
        If Int(NowPage) > Int(TotalPage) Then
            NowPage = TotalPage
        End If

    
    ShowNewJoe = ""
    Dim No
    No=1
    While Not rs.Eof 
         
        If No >= (NowPage-1)*5+1 and No <= NowPage*5 Then
    
        if No Mod 2 = 0 Then
                ShowNewJoe = ShowNewJoe & "<tr bgcolor='#EEEEEE'>"
        Else
                ShowNewJoe = ShowNewJoe & "<tr bgcolor='#FFFFFF'>"
        End If
        
        ShowNewJoe = ShowNewJoe & "<td width='12%'>" & ShowDate(rs("PublishDate")) & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='23%'>" & rs("Service") & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='8%'>" & rs("ProfessionalTitle") & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='18%'>" & rs("Place") & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='12%'>" & rs("OfficePosition") & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='13%'>" & ShowDate(rs("Date")) & "</td>"
        ShowNewJoe = ShowNewJoe & "<td width='14%'><a href='business1_14.asp?Id=" & rs("ID") & "'>查明細</a></td>"
        ShowNewJoe = ShowNewJoe & "</tr>"

        End If
                
        rs.MoveNext
        No = No + 1
        
    WEnd


End Function
'==========================================================================

   Function CheckIdNumber( IdNo )
      ID = UCase(Trim(CStr(IdNo)))
      If Len(ID) <> 10 Then
         CheckIdNumber = false
      ElseIf Asc(ID) < Asc("A") Or Asc(ID) > Asc("Z") Then
         CheckIdNumber = false
      Else
         For I = 1 to 9
            If not IsNumeric(Right(ID, 9)) Then
               CheckIdNumber = false
            End If
         Next
         If I > 9 Then
            x1 = Fix(Map(ID) / 10)
            x2 = Map(ID) Mod 10
            
            d1 = Fix(CLng(Right(ID, 9)) / 100000000)
            d2 = Fix(CLng(Right(ID, 8)) / 10000000)
            d3 = Fix(CLng(Right(ID, 7)) / 1000000)
            d4 = Fix(CLng(Right(ID, 6)) / 100000)
            d5 = Fix(CLng(Right(ID, 5)) / 10000)
            d6 = Fix(CLng(Right(ID, 4)) / 1000)
            d7 = Fix(CLng(Right(ID, 3)) / 100)
            d8 = Fix(CLng(Right(ID, 2)) / 10)
            d9 = Fix(Right(ID, 1))

            Y = x1 + x2*9 + d1*8 + d2*7 + d3*6 + d4*5 + d5*4 + d6*3 + d7*2 + d8 + d9
            If (Y Mod 10) = 0 Then
               CheckIdNumber = true
            Else 
               CheckIdNumber = false            
            End If
         End If
      End If    
   End Function


%>
