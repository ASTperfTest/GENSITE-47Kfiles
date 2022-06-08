Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_knowledge_rss
  Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim Type As String = ""
    Dim TypeName As String = ""
    Dim HeaderCount As String = "20"
    Dim xPostDateStr As String = ""    
    Dim ArticleType As String = ""
    Dim sb As New StringBuilder

    Type = Request.QueryString("type")
    If Type = "A" Then
      TypeName = "知識主題推薦討論"
      ArticleType = "A"
    ElseIf Type = "B" Then
      TypeName = "最新發問"
      ArticleType = "A"
    ElseIf Type = "C" Then
      TypeName = "熱門討論"
      ArticleType = "B"
    ElseIf Type = "D" Then
      TypeName = "專家補充"
      ArticleType = "E"
    End If

    sb.AppendLine("<?xml version=""1.0""  encoding=""utf-8"" ?>")
    sb.AppendLine("<rss version=""2.0"">")
    sb.AppendLine("<channel>")
 
    sb.AppendLine("<title>農業知識入口網-農業知識家</title>")
    sb.AppendLine("<link>http://kmbeta.coa.gov.tw/knowledge/knowledge.aspx</link>")
    sb.AppendLine("<description>農業知識家 - " & TypeName & "</description>")
    sb.AppendLine("<language>zh-tw</language>")
    sb.AppendLine("<lastBuildDate>" & ServerDatetimeGMT(Now()) & "</lastBuildDate>")
    sb.AppendLine("<ttl>" & HeaderCount & "</ttl>")

    If Type = "A" Or Type = "B" Or Type = "C" Or Type = "D" Then

      Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
      Dim sqlString As String = ""
      Dim myConnection As SqlConnection
      Dim myCommand As SqlCommand
      Dim myReader As SqlDataReader
      Dim xURL As String = ConfigurationManager.AppSettings("myURL") & "/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType={1}&CategoryId={2}"

      myConnection = New SqlConnection(ConnString)
      Try
        sqlString = "SELECT TOP " & HeaderCount & " iCUItem, xPostDate, sTitle, xBody, xNewWindow, topCat FROM CuDTGeneric "
        sqlString &= "INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
        sqlString &= "WHERE (iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') "
        If Type = "A" Then
          sqlString &= "AND (CuDTGeneric.xImportant <> 0) ORDER BY xImportant Desc, xPostDate DESC"
        ElseIf Type = "B" Then
          sqlString &= "ORDER BY xPostDate DESC"
        ElseIf Type = "C" Then
          sqlString &= "ORDER BY DiscussCount DESC"
        ElseIf Type = "D" Then
          sqlString &= "AND (KnowledgeForum.HavePros = 'Y') ORDER BY CuDTGeneric.xPostDate DESC"
        End If
   
        myConnection.Open()

        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@ictunit", ConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
        myCommand.Parameters.AddWithValue("@siteid", ConfigurationManager.AppSettings("SiteId"))
        myReader = myCommand.ExecuteReader

        While myReader.Read

          xPostDateStr = ""
          If myReader("xPostDate") IsNot DBNull.Value Then
            xPostDateStr = ServerDatetimeGMT(myReader("xPostDate"))
          End If
          sb.AppendLine("<item iCuItem=""" & myReader("iCuItem") & """ newWindow=""" & myReader("xNewWindow") & """>")
          sb.AppendLine("<title>" & myReader("sTitle") & "</title>")
          sb.Append("<link>" & deAmp(xURL.Replace("{0}", myReader("iCuItem")).Replace("{1}", ArticleType).Replace("{2}", myReader("topCat"))) & "</link>")
          sb.Append("<guid isPermaLink=""false"">" & myReader("iCuItem") & "</guid>")
          sb.Append("<pubDate>" & xPostDateStr & "</pubDate>")
          sb.Append("<description><![CDATA[" & myReader("xBody") & "]]></description>")
          sb.AppendLine("</item>")

        End While

        If Not myReader.IsClosed Then
          myReader.Close()
        End If

      Catch ex As Exception
        Response.Write(ex.ToString)
      Finally
        If myConnection.State = ConnectionState.Open Then
          myConnection.Close()
        End If
      End Try

    End If

    sb.AppendLine("</channel>")
    sb.AppendLine("</rss>")

    Response.Charset = "utf-8"
    Response.ContentType = "text/xml"
    Response.Write(sb.ToString)

  End Sub

  Public Function ServerDatetimeGMT(ByVal dt As DateTime) As String

    Dim DTGMT As DateTime = DateAdd("h", -8, dt)
    Dim WeekDayStr As String = ""
    Select Case Weekday(DTGMT)
      Case 1 : WeekDayStr = "Sun"
      Case 2 : WeekDayStr = "Mon"
      Case 3 : WeekDayStr = "Tue"
      Case 4 : WeekDayStr = "Wed"
      Case 5 : WeekDayStr = "Thu"
      Case 6 : WeekDayStr = "Fri"
      Case 7 : WeekDayStr = "Sat"
    End Select
    Dim MonthStr As String = ""
    Select Case Month(DTGMT)
      Case 1 : MonthStr = "Jan"
      Case 2 : MonthStr = "Feb"
      Case 3 : MonthStr = "Mar"
      Case 4 : MonthStr = "Apr"
      Case 5 : MonthStr = "May"
      Case 6 : MonthStr = "Jun"
      Case 7 : MonthStr = "Jul"
      Case 8 : MonthStr = "Aug"
      Case 9 : MonthStr = "Sep"
      Case 10 : MonthStr = "Oct"
      Case 11 : MonthStr = "Nov"
      Case 12 : MonthStr = "Dec"
    End Select
    Dim xhour As String = Right("00" + CStr(Hour(DTGMT)), 2)
    Dim xminute As String = Right("00" + CStr(Minute(DTGMT)), 2)
    Dim xsecond As String = Right("00" + CStr(Second(DTGMT)), 2)
    If Len(dt) = 0 Then
      Return ""
    Else
      Return WeekDayStr + ", " + Right("00" + CStr(Day(DTGMT)), 2) + " " + MonthStr + " " + CStr(Year(DTGMT)) + " " + xhour + ":" + xminute + ":" + xsecond + " GMT"
    End If

  End Function

  Public Function deAmp(ByVal tempstr As String)

    If tempstr = "" Or tempstr = Nothing Then
      Return ""
    Else
      Return tempstr.Replace("&", "&amp;")
    End If

  End Function

End Class
