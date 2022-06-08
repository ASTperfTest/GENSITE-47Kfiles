Imports System.Web.Configuration
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.Sql
Imports System.Data.SqlClient

Partial Class xdSP
    Inherits System.Web.UI.Page

    Dim mp As String
    Dim connstringName As String
    Dim debug As Boolean = False
    Dim xdURL As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            connstringName = WebConfigurationManager.AppSettings("connstringName")
            'Dim mySiteID = ConfigurationSettings.AppSettings("mySiteID")
            'Dim XMLOutput As New xdSPXML

            mp = Request("mp")
            xdURL = Request("xdURL")

            If Request("debug") <> "" Then
                debug = True
            End If

            If xdURL = "" Then
                'Response.Write(Resources.Resource.ErrorXDURL)
            ElseIf mp = "" Then
                'Response.Write(Resources.Resource.ErrorMP)
            Else
                Dim xml As String
                xml = xdSPXMLGen()
                Response.Write(xml)
            End If
        Catch ex As Exception
            Response.Write(ex)
            'Throw ex
        Finally
        End Try
        Response.End()
    End Sub

    Public Function xdSPXMLGen() As String
        Dim xmlsb As StringBuilder = New StringBuilder

        xmlsb.Append("<?xml version=""1.0""  encoding=""utf-8"" ?>" & vbCrLf & "<hpMain>")

        Dim writer As New StringWriter
        Dim pHTML As String
        Server.Execute(xdURL, writer)
        pHTML = "<pHTML><![CDATA[" + writer.ToString() + "]]></pHTML>" & vbCrLf

        'Dim retString As String

        Dim htPageDom As XmlDocument = New XmlDocument

        Dim LoadXML As String = System.IO.Path.Combine(WebConfigurationManager.AppSettings("gipdsd"), "xdmp" & mp & ".xml")
        htPageDom.Load(LoadXML)
        Dim refModel As XmlNode = htPageDom.SelectSingleNode("//MpDataSet")
        Dim xRootID As String = CommFunc.nullText(refModel.SelectSingleNode("MenuTree"))
        xmlsb.Append("<MenuTitle>" & CommFunc.nullText(refModel.SelectSingleNode("MenuTitle")) & "</MenuTitle>")
        xmlsb.Append("<myStyle>" & CommFunc.nullText(refModel.SelectSingleNode("MpStyle")) & "</myStyle>")
        xmlsb.Append("<mp>" & mp & "</mp>")

        Dim redirectUrl As String = Request.QueryString("redirectUrl")
        xmlsb.Append("<redirectUrl><![CDATA[" & redirectUrl & "]]></redirectUrl>")

        xmlsb.Append("<login>")
        Dim memName As String = ""
		Dim memnickname As String = ""
        If Request("memID") <> "" Then		
		
            Dim sql As String = "SELECT realname, ISNULL(nickname, '') AS nickname, ISNULL(logincount, 0) AS logincount  FROM Member where account = '" & Request("memID") & "'"
            Dim myconn As New SqlConnection(WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString)
            Try
                myconn.Open()
                Dim mycomm As New SqlCommand(sql, myconn)
                Dim myReader As SqlDataReader
                myReader = mycomm.ExecuteReader
                If myReader.Read() Then
                    xmlsb.Append("<status>true</status>")
                    memName = myReader.Item("realname")
					memnickname = myReader.Item("nickname")
                Else
                    xmlsb.Append("<status>false</status>")
                End If

            Catch ex As Exception
            Finally
                If myconn.State = Data.ConnectionState.Open Then
                    myconn.Close()
                End If
            End Try
        Else
            xmlsb.Append("<status>false</status>")
        End If

        xmlsb.Append("<memID>" & Request("memID") & "</memID>")
        xmlsb.Append("<memName><![CDATA[" & memName & "]]></memName>")
		xmlsb.Append("<memNickName><![CDATA[" & memnickname & "]]></memNickName>")
        xmlsb.Append("<gstyle>" & Request("gstyle") & "</gstyle>")
        xmlsb.Append("</login>")

    xmlsb.Append(GetWebActivity())
	
	xmlsb.Append(GetCurrentRead())
	
        Dim NowTime As Date
        Dim NowHour As Integer
        NowTime = Now
        NowHour = Integer.Parse(NowTime.Hour)
        If NowHour >= 6 And NowHour < 12 Then
            xmlsb.Append("<NowTime>AM</NowTime>")
        ElseIf NowHour >= 12 And NowHour < 18 Then
            xmlsb.Append("<NowTime>PM</NowTime>")
        Else
            xmlsb.Append("<NowTime>NT</NowTime>")
        End If
        '----處理xdSP	
        xmlsb.Append(pHTML)

        Dim myTreeNode As Integer = 0
        Dim upParent As Integer = 0
        Dim xmlDataSet As XmlNode
		
		'使用 ProcessXDataSet Function 的SqlCondition參數傳日出日落的額外判斷條件 2009.08.18 By Ivy
		Dim sqlConditionForSun As String = ""
		If Request("memID") = "" Then
			sqlConditionForSun = " AND xKeyword in('台北','') "
		Else
			sqlConditionForSun = " AND xKeyword =CASE (SELECT keyword FROM member WHERE account='" & Request("memID") & "') WHEN '' THEN '台北' ELSE (SELECT keyword FROM member WHERE account='" & Request("memID") & "') END"
		End If
		
        For Each xmlDataSet In refModel.SelectNodes("DataSet[ContentData='Y']")
            'xmlsb.Append(vbCrLf & "<now3>" & Now & "</now3>" & vbCrLf)		
            xmlsb.Append(xDataSet.ProcessXDataSet(xmlDataSet, mp, myTreeNode, debug, "", sqlConditionForSun, ""))'modify by ivy
            'xmlsb.Append(vbCrLf & "<now4>" & Now & "</now4>" & vbCrLf)
        Next
	
		
		
		
		

        Dim topParent As Integer = 0
        '----導覽目錄樹選單
        xmlsb.Append(Menus.x1Menus(refModel, myTreeNode, upParent, topParent, mp))
        'xmlsb.Append("<Hit>" & CommFunc.Counter("1", "0", Server.MapPath("counter.txt"), Application) & "</Hit>")
				xmlsb.Append("<etoday>" & Now() & "</etoday>")
        xmlsb.Append("<today>民國" & Now.Year - 1911 & "年" & Now.Month & "月" & Now.Day & "日</today>")
        xmlsb.Append("</hpMain>")
        Return xmlsb.ToString
        
  End Function

  Function GetWebActivity() As String
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings.Item("ConnString").ToString()
    Dim MemberConn As SqlConnection = New SqlConnection(ConnString)
    Dim MemberComm As SqlCommand
    Dim MemberReader As SqlDataReader
    Dim SqlString As String
    Dim str As String = ""

    str = "<webActivity>"

    Try
      SqlString = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1'"
      MemberConn.Open()
      MemberComm = New SqlCommand(SqlString, MemberConn)

      MemberReader = MemberComm.ExecuteReader

      If MemberReader.Read() Then
        str &= "<active>true</active>"
        str &= "<id>" & MemberReader("m011_subjectid") & "</id>"
      Else
        str &= "<active>false</active>"
        str &= "<id></id>"
      End If

    Catch ex As Exception      
    Finally
      If MemberConn.State = Data.ConnectionState.Open Then
        MemberConn.Close()
      End If
    End Try
    str &= "</webActivity>"

    Return str

  End Function
  
  Function GetCurrentRead() As String
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings.Item("ConnString").ToString()
    Dim ReadConn As SqlConnection = New SqlConnection(ConnString)
    Dim ReadComm As SqlCommand
    Dim ReadReader As SqlDataReader
    Dim SqlString As String
    Dim str As String = ""
	Dim i as integer
	Dim CurrentReadList() As String

        str = "<newest>"

    Try
		CurrentReadList = split(Request("CurrentRead"),",")
		for i = 0 to 4
			SqlString = "SELECT sTitle,iCUItem,topCat FROM CuDTGeneric WHERE iCUItem = " & CurrentReadList(i)
			ReadConn.Open()
			ReadComm = New SqlCommand(SqlString, ReadConn)
			ReadReader = ReadComm.ExecuteReader
			If ReadReader.Read() Then
				str &= "<article>"
				str &= "<title><![CDATA[" & ReadReader("sTitle") & "]]></title>"
				str &= "<url><![CDATA[knowledge_cp.aspx?ArticleId=" & ReadReader("iCUItem") & "&amp;ArticleType=A&amp;CategoryId=" & Trim(ReadReader("topCat")) & "]]></url>"
				str &= "</article>"
			End if
			If ReadConn.State = Data.ConnectionState.Open Then
				ReadConn.Close()
			End If
		Next
	  

    Catch ex As Exception      
    Finally
      If ReadConn.State = Data.ConnectionState.Open Then
        ReadConn.Close()
      End If
    End Try
    str &= "</newest>"

    Return str

  End Function
End Class
