Imports System.Web.Configuration
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient

Partial Class xdlogin
    Inherits System.Web.UI.Page

  Dim connstringName As String
  Dim debug As Boolean = False
  Dim userid As String = ""
  Dim password As String = ""
  Dim mp As String

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      connstringName = WebConfigurationManager.AppSettings("connstringName")

      mp = Request("mp")
      userid = Request("account2")
      password = Request("passwd2")

      If Request("debug") <> "" Then
        debug = True
      End If

      If mp = "" Then
        'Response.Write(Resources.Resource.ErrorMP)
      Else
        Dim xml As String
        xml = xdXMLGen()
        Response.Write(xml)
      End If
    Catch ex As Exception
      Response.Write(ex)
      'Throw ex
    Finally
    End Try
    Response.End()
  End Sub

  Public Function xdXMLGen() As String

    Dim xmlsb As StringBuilder = New StringBuilder
    xmlsb.Append("<?xml version=""1.0""  encoding=""utf-8"" ?>" & vbCrLf & "<hpMain>")
    Dim writer As New StringWriter

    'Dim retString As String
    Dim debug As Boolean = True
    Dim htPageDom As XmlDocument = New XmlDocument

    Dim LoadXML As String = WebConfigurationManager.AppSettings("gipdsd") & "\xdmp" & mp & ".xml"
    htPageDom.Load(LoadXML)
    Dim refModel As XmlNode = htPageDom.SelectSingleNode("//MpDataSet")
    Dim xRootID As String = CommFunc.nullText(refModel.SelectSingleNode("MenuTree"))
    xmlsb.Append("<MenuTitle>" & CommFunc.nullText(refModel.SelectSingleNode("MenuTitle")) & "</MenuTitle>")
    xmlsb.Append("<myStyle>" & CommFunc.nullText(refModel.SelectSingleNode("MpStyle")) & "</myStyle>")
    xmlsb.Append("<mp>" & mp & "</mp>")

    xmlsb.Append("<login>" & Login() & "</login>")

    Dim myTreeNode As Integer = 0
    Dim upParent As Integer = 0

    Dim xmlDataSet As XmlNode
    For Each xmlDataSet In refModel.SelectNodes("DataSet[ContentData='Y']")
      'xmlsb.Append(vbCrLf & "<now3>" & Now & "</now3>" & vbCrLf)
      xmlsb.Append(xDataSet.ProcessXDataSet(xmlDataSet, mp, myTreeNode, debug, "", "", ""))
      'xmlsb.Append(vbCrLf & "<now4>" & Now & "</now4>" & vbCrLf)
    Next

    Dim topParent As Integer = 0

    '----導覽目錄樹選單
    xmlsb.Append(Menus.x1Menus(refModel, myTreeNode, upParent, topParent, mp))

    xmlsb.Append("</hpMain>")
    Return xmlsb.ToString

  End Function

  Public Function Login() As String

    Dim xmlsb As StringBuilder = New StringBuilder
    Dim bLogin As Boolean = False
    Dim memName As String = ""
    Dim gstyle As String = ""
    Dim memNickName As String = ""
    Dim memLoginCount As Integer = 0
    Dim Type As String = "0"
    Dim sql As String = ""
    Dim bolexpert As Boolean = True, bolscholar As Boolean = True, bolmember As Boolean = True
    Dim dbconn As SqlConnection
	
	Dim birthday As String = ""
	Dim validBirthday As Boolean = False
	Dim mcode As String = ""

    dbconn = New SqlConnection
    dbconn.ConnectionString = WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString

    Try
      dbconn.Open()

      sql = "SELECT realname,ISNULL(birthday, '20000101') AS birthday, ISNULL(nickname, '') AS nickname, ISNULL(logincount, 0) AS logincount, "
      sql &= "ISNULL(status, '') AS status, ISNULL(actor, '') AS actor, ISNULL(id_type1, '') AS id_type1, "
      sql &= "ISNULL(id_type2, '') AS id_type2, ISNULL(id_type3, '') AS id_type3, ISNULL(scholarValidate, '') AS scholarValidate "
      sql &= ",ISNULL(mcode, '') AS mcode "
      sql &= "FROM Member where account = @UserID and passwd = @Password "

      Dim oCommand As SqlCommand = New SqlCommand(sql, dbconn)
      oCommand.Parameters.AddWithValue("@UserID", userid)
      oCommand.Parameters.AddWithValue("@Password", password)

      Dim objReader As SqlDataReader = oCommand.ExecuteReader()

      If objReader.Read Then

        '---2008/11/12---
        '---一般會員---id_type1=1        
        '---學者會員---id_type2=1
        '---專家會員---id_type3=1
        '---1. 檢查此帳號是否停權, 若停權, login = false, type = 2---
        '---2. 若是專家會員
        '---3. 若是學者會員, 檢查審核是否通過. 若通過, 以學者身份登入, 若沒通過, 以一般會員檢查.
        '---4. 若是一般會員

        '---1. 檢查此帳號是否停權, 若停權, login = false, type = 2---
        If objReader("status") = "Y" or objReader("status") = "" Then

          '---2. 若是專家會員
          If bolexpert Then
            If objReader("id_type3") = "1" Then
              bLogin = True
              gstyle = "005"
              bolscholar = False
              bolmember = False            
            End If
          End If

          '---3. 若是學者會員, 檢查審核是否通過. 若通過, 以學者身份登入, 若沒通過, 以一般會員檢查.
          If bolscholar Then
            If objReader("id_type2") = "1" Then
              If objReader("scholarValidate") = "Y" Then
                bLogin = True
                gstyle = "003"
                bolmember = False
              End If
            End If
          End If

          '---4. 若是一般會員
          If bolmember Then
            If objReader("id_type1") = "1" Then
              bLogin = True
              gstyle = "002"
            End If
          End If

          If bLogin Then
            '---給值---
            memName = objReader("realname").ToString.Trim
            memNickName = objReader("nickname").ToString.Trim
            memLoginCount = objReader("logincount")   
			birthday = objReader("birthday") 
			validBirthday = CheckBirthdayRight(birthday)
			mcode = objReader("mcode") 
          End If

          'If objReader("actor").ToString.Trim = "1" Or objReader("actor").ToString.Trim = "2" Or objReader("actor").ToString.Trim = "3" Then '---學者---
          '  If objReader("status").ToString.Trim = "Y" Then '---審核通過---
          '    gstyle = "003"
          '    bLogin = True
          '  Else
          '    '---審核不通過---
          '    gstyle = ""
          '    bLogin = True
          '    Type = "2"
          '  End If
          'ElseIf objReader("actor").ToString.Trim = "5" Then '---專家---
          '  bLogin = True
          '  gstyle = "005"
          'ElseIf objReader("actor").ToString.Trim = "0" Or objReader("actor").ToString.Trim = "" Or objReader("actor") Is DBNull.Value Then '---一般---
          '  bLogin = True
          '  gstyle = "002"
          'End If

        Else
          '---帳號被停權---login = false, type = 2---
          Type = "2"
        End If
      Else
        '---帳號或密碼錯誤---
        Type = "1"
      End If
      objReader.Close()

      If bLogin Then
        '---更新Login count---
        sql = "UPDATE Member SET logincount = ISNULL(logincount, 0) + 1 WHERE account = @UserID "
        oCommand = New SqlCommand(sql, dbconn)
        oCommand.Parameters.AddWithValue("@UserID", userid)
        oCommand.ExecuteNonQuery()
      End If

    Catch ex As Exception
      'Response.Write(ex)      
    Finally
      dbconn.Close()
    End Try

    xmlsb.Append("<status>" & bLogin & "</status>")
    xmlsb.Append("<memID><![CDATA[" & userid & "]]></memID>")
    xmlsb.Append("<memName><![CDATA[" & memName & "]]></memName>")
    xmlsb.Append("<memNickName><![CDATA[" & memNickName & "]]></memNickName>")
    xmlsb.Append("<memLoginCount>" & memLoginCount & "</memLoginCount>")
    xmlsb.Append("<gstyle>" & gstyle & "</gstyle>")
    xmlsb.Append("<type>" & Type & "</type>")
	'xmlsb.Append("<birthday>" & birthday & "</birthday>")
	xmlsb.Append("<validBirthday>" & validBirthday & "</validBirthday>")
	xmlsb.Append("<mcode>" & mcode & "</mcode>")

    Return xmlsb.ToString

  End Function
  
  '檢查出生日期是否為有效日期
  '無效情形:NULL,空字串,不滿8碼  
  Public Function CheckBirthdayRight(ByVal s As String) As Boolean
  
    Dim boolDate As Boolean = False

	If s.Length <> 8 Then
		boolDate = False            
	Else
		'將8碼的生日日期替換成格式yyyy/mm/dd
		s = s.Insert(6, "/")
		s = s.Insert(4, "/")

		boolDate = IsDate(s)
		
	End If       

	Return boolDate
  
  End Function
    
End Class
