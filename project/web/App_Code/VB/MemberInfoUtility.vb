Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class MemberInfoUtility

  Public Shared Function ReplaceMemberName(ByVal account As String, ByVal name As String, ByVal gradeLevel As Integer) As Array
    Dim strTemp As String = name
    Dim memberInfoTag(1) As String
    Dim memberInfoTableHtml As String = ""
	
	'判斷是否為高手極以上會員，是則提供即時顯示特殊資訊功能
	If gradeLevel > 2 Then
		memberInfoTableHtml &= WriteMemberInfoTag(account, name)
		strTemp = Replace(strTemp, strTemp, "<a title='" &GetPersonalValue(account) &"' id=""account_" & account & """ onclick='controlDiv(""account_" & account & """,this);AdjustPhoto(""" & account & """);' class=""pediaword"" style=""color:purple;""><span>" & strTemp + "</span></a>", 1, 1)
	Else
		strTemp = "<a title='" &GetPersonalValue(account) &"'> " & name & "</a>"
	End If
    
    memberInfoTag(1) = memberInfoTableHtml
    memberInfoTag(0) = strTemp
    Return memberInfoTag

  End Function

  Public Shared Function WriteMemberInfoTag(ByVal account As String, ByVal name As String) As String
    Dim sb1 As String = GetMemberInfoTableHtml(account, name)
    Dim htmtemp As String = ""
    htmtemp &= "<Table id='inc_account_" & account & "' class='member'><tbody>"
    htmtemp &= "<tr><td align=""right""><div class=""aasclose"" onclick='closeDiv(""account_" & account & """);'>X</div></td></tr>"
    htmtemp &= "<tr><td>" & sb1 & "</td></tr></tbody></table>"
    Return htmtemp
  End Function

  Private Shared Function GetPhotoSrc(ByVal account As String, ByVal photo As String) As String

	Dim filePath As String = "http://kmwebsys.coa.gov.tw/public/Profile/"
	if photo = "" Then
	filePath = filePath & "default.png"
	Else
	filePath = filePath & account & "\" &  photo
	End If

    Return filePath

  End Function

  Public Shared Function GetMemberInfoTableHtml(ByVal account As String, ByVal name As String) As String
    Dim sqlString As String = ""
    Dim tableHtm As String = ""
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim cnn As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    cnn = New SqlConnection(ConnString)
    Try
      sqlString = "  SELECT photo,introduce,contact FROM Member WHERE account=@account"

      cnn.Open()
      myCommand = New SqlCommand(sqlString, cnn)
      myCommand.Parameters.AddWithValue("@account", account)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
		tableHtm &= "<table>"
		tableHtm &= "<tbody>"
		tableHtm &= "<tr>"
		tableHtm &= "<td valign='top' align='middle' >"
		tableHtm &= "<div >"
		tableHtm &= "<img id='memberPhoto_" & account & "'  alt='" & name &"' src=" & GetPhotoSrc(account, Convert.ToString(myReader("photo"))) & " align='center'  border='0'  > "
		tableHtm &= "</div>"
		tableHtm &= "</td>"
		tableHtm &= "<td>"
		tableHtm &= "<table width='280px' cellspacing='1' cellpadding='3' border='0' bgcolor='#f0f0f0' >"
		tableHtm &= "<tbody>"
		tableHtm &= "<tr >"
		tableHtm &= "<td bgcolor='#ffffff' align='center'>" & name & "</td>"
		tableHtm &= "</tr>"
		tableHtm &= "</tbody>"
		tableHtm &= "</table>"
		tableHtm &= "<table width='280px' cellspacing='1' cellpadding='3' border='0' bgcolor='#f0f0f0' >"
		tableHtm &= "<tbody>"
		tableHtm &= "<tr >"
		tableHtm &= "<td bgcolor='#f1f6fc' align='right' width='60'>自我介紹</td>"
		tableHtm &= "<td bgcolor='#ffffff'>" 
		tableHtm &= IIf(Convert.ToString(myReader("introduce")) = "", "無", myReader("introduce")) & "</td>"
		tableHtm &= "</tr>"
		tableHtm &= "<tr>"
		tableHtm &= "<td bgcolor='#f1f6fc' align='right' width='60'>網址</td>"
		tableHtm &= "<td bgcolor='#ffffff'>" & IIf(Convert.ToString(myReader("contact")) = "", "無", myReader("contact")) & "</td>"
		tableHtm &= "</tr>"
		tableHtm &= "</tbody>"
		tableHtm &= "</table>"
		tableHtm &= "</td>"
		tableHtm &= "</tr>"	
		tableHtm &= "</tbody>"
		tableHtm &= "</table>"

      End If

      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
    End Try
    Return tableHtm
  End Function
  
  '取得個人評價
  Private Shared Function GetPersonalValue(ByVal account As String)As String
  	
	Dim str As String = ""
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "SELECT CAST(SUM (GradeCount / GradePersonCount) / COUNT(*) AS DECIMAL(5,2)) AVGSTART " 
	  sqlString &= "FROM ( " 
	  sqlString &= "SELECT CAST(KnowledgeForum.GradeCount AS DECIMAL(5,2))GradeCount, CAST(KnowledgeForum.GradePersonCount AS DECIMAL(5,2)) GradePersonCount " 
	  sqlString &= "FROM CuDTGeneric AS CuDTGeneric_1 " 
	  sqlString &= "INNER JOIN CuDTGeneric " 
	  sqlString &= "INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem " 
	  sqlString &= "INNER JOIN KnowledgeForum AS KnowledgeForum_1 " 
	  sqlString &= "ON CuDTGeneric.iCUItem = KnowledgeForum_1.ParentIcuitem " 
	  sqlString &= "ON CuDTGeneric_1.iCUItem = KnowledgeForum_1.gicuitem " 
	  sqlString &= "INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
	  sqlString &= "WHERE (CuDTGeneric_1.iCTUnit = @ictunit) " 
	  sqlString &= "AND (CuDTGeneric_1.iEditor = @account) " 
	  sqlString &= "AND (CodeMain.codeMetaID = 'KnowledgeType') " 
	  sqlString &= "AND (CuDTGeneric_1.siteId = @siteid) " 
	  sqlString &= "AND (KnowledgeForum_1.Status = 'N') " 
   	  sqlString &= "AND (KnowledgeForum.Status = 'N') " 
   	  sqlString &= "AND (KnowledgeForum.GradePersonCount >0) " 
	  sqlString &= ") TA  "
      
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@account", account)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
      myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
      myReader = myCommand.ExecuteReader()
	
	  If myReader.Read Then
		str = Convert.ToString(myReader("AVGSTART"))
	  End If
     
      If Not myReader.IsClosed Then
        myReader.Close()
      End If

      myCommand.Dispose()

    Catch ex As Exception
      
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return "平均評價：" & str & "顆星(滿分5顆星)"
	
  End Function

End Class