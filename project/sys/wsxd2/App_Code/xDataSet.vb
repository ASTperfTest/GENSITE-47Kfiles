Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Web.Configuration
Imports System.Data.SqlClient

Public Class xDataSet
    Shared Function ProcessXDataSet(ByVal xDataSet As XmlNode, ByVal mp As String, ByVal ctNode As String, ByVal debug As Boolean, ByVal KMLikeStr As String, ByVal SqlCondition As String, ByVal SqlOrderBy As String) As String
        
        Dim xmlsb As StringBuilder = New StringBuilder
	    	Dim connstringName As String = WebConfigurationManager.AppSettings("connstringName")

        '----GIP Node	
        'xmlsb.Append(vbCrLf & "<now5>" & Now & "</now5>" & vbCrLf)
        Dim xTreeNode As String = CommFunc.nullText(xDataSet.SelectSingleNode("DataNode"))
        If xTreeNode = "" Then xTreeNode = 0
        Dim sql As String = "SELECT t.*, b.sBaseTableName " _
         & " FROM CatTreeNode AS t LEFT JOIN CtUnit AS u ON u.CtUnitID=t.CtUnitID " _
         & " LEFT JOIN BaseDSD AS b ON b.iBaseDSD=u.iBaseDSD " _
         & " WHERE CtNodeID IN (" & xTreeNode & ")"
        If debug <> True Then xmlsb.Append("<xNodesql><![CDATA[" & sql & "]]></xNodesql>")

        Dim rssurl As String = ""

        Dim dbconn As SqlConnection
				
        dbconn = New SqlConnection
        dbconn.ConnectionString = WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString
        'Return "<xNodesql><![CDATA[" & WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString & "]]></xNodesql>"
        Try         
            dbconn.Open()                         
            Dim oCommand As SqlCommand = New SqlCommand(sql, dbconn)            
            Dim objReader As SqlDataReader = oCommand.ExecuteReader()            
            'xmlsb.Append(vbCrLf & "<now6>" & Now & "</now6>" & vbCrLf)


            If objReader.Read Then
                Dim xDataNode As String = objReader("CtUnitID")
                Dim xBaseTableName As String = LCase(objReader("sBaseTableName"))
                Dim dCondition As String = Trim(objReader("dCondition") & " ")
                xmlsb.Append("<" & CommFunc.nullText(xDataSet.SelectSingleNode("DataLable")) _
                  & " xNode=""" & objReader("CtNodeID") & """ xUnit=""" & xDataNode & """>")

                xmlsb.Append("<ContentStyle>" & CommFunc.nullText(xDataSet.SelectSingleNode("ContentStyle")) & "</ContentStyle>")
                xmlsb.Append("<DataSrc>" & CommFunc.nullText(xDataSet.SelectSingleNode("DataSrc")) & "</DataSrc>")
                xmlsb.Append("<PicWidth>" & CommFunc.nullText(xDataSet.SelectSingleNode("PicWidth")) & "</PicWidth>")
                xmlsb.Append("<PicHeight>" & CommFunc.nullText(xDataSet.SelectSingleNode("PicHeight")) & "</PicHeight>")

                Dim ContentStyle As String = CommFunc.nullText(xDataSet.SelectSingleNode("ContentStyle"))

                xmlsb.Append("<Caption>" & CommFunc.nullText(xDataSet.SelectSingleNode("DataRemark")) & "</Caption>")
                Dim MprefDataBlock As String = CommFunc.nullText(xDataSet.SelectSingleNode("refDataBlock"))

                '	MprefHeaderrandom = nullText(xDataSet.selectSingleNode("refIsrandom"))
                '	MprefSqlCondition = nullText(xDataSet.selectSingleNode("refSqlCondition"))
                '	MprefHeaderCount = nullText(xDataSet.selectSingleNode("refHeaderCount"))
                '	MprefSqlOrderBy = nullText(xDataSet.selectSingleNode("refSqlOrderBy"))

                Dim ContentLength As String = CommFunc.nullText(xDataSet.SelectSingleNode("ContentLength"))
                If ContentLength = "" Then ContentLength = "120"


                Dim HeaderCount As String = CommFunc.nullText(xDataSet.SelectSingleNode("SqlTop"))
                If debug = True Then xmlsb.Append("<SqlTop>" & HeaderCount & "</SqlTop>")

                Dim Headerrandom As String = CommFunc.nullText(xDataSet.SelectSingleNode("Israndom"))
                If HeaderCount <> "" Then HeaderCount = "TOP " & HeaderCount

                sql = "SELECT " & HeaderCount & " htx.*, xr1.deptName, xr1.edeptName, u.CtUnitName, n.CtNodeID "
                If Headerrandom = "Y" Then sql = sql & ", RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3771) " _
                  & " * RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3) as ra "
                '	if xBaseTableName<>"" then	sql = sql & "JOIN " & xBaseTableName & " AS ghtx ON htx.iCuItem=ghtx.giCuItem"
                sql = sql & " FROM CuDTGeneric AS htx " _
                 & " JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
                 & " JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit" _
                 & " LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept" _
                 & " WHERE htx.fCTUPublic='Y' " _
                 & " AND (htx.avEnd is NULL OR htx.avEnd >=" & CommFunc.pkStr(Date.Now.Date, ")") _
                 & " AND (htx.avBegin is NULL OR htx.avBegin <=" & CommFunc.pkStr(Date.Now.Date, ")")
                '		& " LEFT JOIN CuDTwImg AS img ON img.gicuitem=htx.iCuItem" _
                '	if xDataNode<>"" then		sql = sql & " AND iCtUnit=" & xDataNode
                If xTreeNode <> "" Then sql = sql & " AND n.CtNodeID IN (" & xTreeNode & ")"
				
				' 首頁日出日落預設顯示台北區域，使用者登入後，顯示該使用者的所在區域 2009.08.19 By Ivy
				Dim sunsetSqlCondition As String = SqlCondition
				
                SqlCondition = CommFunc.nullText(xDataSet.SelectSingleNode("SqlCondition"))
                If SqlCondition <> "" Then sql = sql & " AND " & SqlCondition
				'2009.08.19 by ivy
				If xTreeNode=1572 then	sql = sql & sunsetSqlCondition
				'end by ivy
                If dCondition <> "" And Not dCondition Is DBNull.Value Then _
                 sql = sql & " AND " & Replace(dCondition, "CuDTGeneric", "htx")

                SqlOrderBy = CommFunc.nullText(xDataSet.SelectSingleNode("SqlOrderBy"))
                If SqlOrderBy <> "" Then sql = sql & " ORDER BY " & SqlOrderBy
                '	xtn = split(xTreeNode, ",")
                '	xTreeNode = xtn(0)
                If debug = True Then
                    xmlsb.Append("<xDataSetsql><![CDATA[" & sql & "]]></xDataSetsql>")
                End If

                Dim dbconn1 As SqlConnection
                dbconn1 = New SqlConnection
                dbconn1.ConnectionString = WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString

                Try
                    dbconn1.Open()
                    Dim oCommand1 As SqlCommand = New SqlCommand(sql, dbconn1)
                    Dim objReader1 As SqlDataReader = oCommand1.ExecuteReader()
                    'xmlsb.Append(vbCrLf & "<now10>" & Now & "</now10>" & vbCrLf)
                    While objReader1.Read
                        'xURL = "ct.asp?xItem=" & RS("iCuItem") & "&amp;ctNode=" & xTreeNode
                        Dim xURL As String = "?a=ct&amp;xItem=" & objReader1("iCuItem") & "&amp;ctNode=" & xTreeNode & "&amp;mp=" & mp
                        If objReader1("ibaseDSD") = 2 And Not objReader1("xURL") Is DBNull.Value Then xURL = CommFunc.deAmp(objReader1("xURL"))
                        If xBaseTableName = "adrotate" And Not objReader1("xURL") Is DBNull.Value Then xURL = CommFunc.deAmp(objReader1("xURL"))
                        If objReader1("showType") = 2 And Not objReader1("xURL") Is DBNull.Value Then xURL = CommFunc.deAmp(objReader1("xURL"))
                        If objReader1("showType") = 3 And Not objReader1("fileDownLoad") Is DBNull.Value Then
                            xURL = "public/Data/" & objReader1("fileDownLoad")
                        End If
                        If objReader1("showType") = 5 Then xURL = "content.asp?cuItem=" & objReader1("refID")

                        xmlsb.Append("<Article iCuItem=""" & objReader1("iCuItem") & """ newWindow=""" & objReader1("xNewWindow") & """>")
                        xmlsb.Append("<Caption>" & objReader1("sTitle") & "</Caption>")

                        If Not objReader1("xBody") Is DBNull.Value Then
                            Dim xBody As String = CommFunc.deHTML(objReader1("xBody"))
                            xBody = Left(xBody, CInt(ContentLength))
                            If Len(xBody) > CInt(ContentLength) Then xBody = Left(xBody, CInt(ContentLength))
                            xmlsb.Append("<Content><![CDATA[" & xBody & "]]></Content>")
                            xmlsb.Append("<ContentAll><![CDATA[" & objReader1("xBody") & "]]></ContentAll>")
                        Else
                            xmlsb.Append("<Content></Content>")
                            xmlsb.Append("<ContentAll></ContentAll>")
                        End If

                        If Not objReader1("xabstract") Is DBNull.Value Then
                            xmlsb.Append("<Abstract><![CDATA[" & objReader1("xabstract") & "]]></Abstract>")
                        End If
                   
                        xmlsb.Append("<PostDate>" & objReader1("xPostDate") & "</PostDate>")
                        xmlsb.Append("<DeptName>" & objReader1("deptName") & "</DeptName>")
                        xmlsb.Append("<TopCat>" & objReader1("TopCat") & "</TopCat>")
                        xmlsb.Append("<EDeptName>" & objReader1("edeptName") & "</EDeptName>")
                        xmlsb.Append("<Group>" & objReader1("vGroup") & "</Group>")
                        xmlsb.Append("<CtUnitName>" & CommFunc.deAmp(objReader1("CtUnitName")) & "</CtUnitName>")
                        
                        'xmlsb.Append("<xURL><![CDATA[" & CommFunc.deAmp(xURL) & "]]></xURL>")
			xmlsb.Append("<xURL>" & xURL & "</xURL>")
                        '---2006/10/23-sql中沒出現---
                        'xmlsb.Append("<hitPageCount>" & objReader1("hitPageCount") & "</hitPageCount>")
                        If Not objReader1("xImgFile") Is DBNull.Value Then
                            xmlsb.Append("<xImgFile>public/data/" & objReader1("xImgFile") & "</xImgFile>")
                        End If

                        If MprefDataBlock <> "" Then
                            Dim xsql As String = "SELECT Top 1 htx.* "
                            xsql = xsql & ", RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3771) " _
                             & " * RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3) as ra "
                            xsql = xsql & " FROM CuDTGeneric AS htx" _
                             & " WHERE htx.fCTUPublic='Y' AND ictunit=" & MprefDataBlock _
                             & " AND (htx.avEnd is NULL OR htx.avEnd >=" & CommFunc.pkStr(Date.Now.Date, ")") _
                             & " AND (htx.avBegin is NULL OR htx.avBegin <=" & CommFunc.pkStr(Date.Now.Date, ")")
                            xsql = xsql & " ORDER BY ra"

                            Dim dbconn2 As SqlConnection
                            dbconn2 = New SqlConnection
                            dbconn2.ConnectionString = WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString
                            Try
                                dbconn2.Open()
                                Dim oCommand2 As SqlCommand = New SqlCommand(xsql, dbconn2)
                                Dim objReader2 As SqlDataReader = oCommand2.ExecuteReader()
                                While objReader2.Read
                                    xmlsb.Append("<refDataBlock iCuItem=""" & objReader2("iCuItem") & """ newWindow=""" & objReader2("xNewWindow") & """>")
                                    xmlsb.Append("<Caption>" & CommFunc.deAmp(objReader2("sTitle")) & "</Caption>")
                                    If Not objReader2 Is DBNull.Value Then
                                        xmlsb.Append("<xImgFile>public/Data/" & objReader2("xImgFile") & "</xImgFile>")
                                    End If
                                    xmlsb.Append("</refDataBlock>")
                                End While
                            Catch ex As Exception
                                Throw ex
                                'xmlsb.Append(vbCrLf & "<error>" & vbCrLf)
                                'xmlsb.Append(ex)
                                'xmlsb.Append(vbCrLf & "</error>" & vbCrLf)
                                'xmlsb.Append(vbCrLf & "<sql>" & vbCrLf)
                                'xmlsb.Append(sql)
                                'xmlsb.Append(vbCrLf & "</sql>" & vbCrLf)
                            Finally
                                dbconn2.Close()
                            End Try
                        End If

                        xmlsb.Append("</Article>")
                    End While
                    'xmlsb.Append(vbCrLf & "<now11>" & Now & "</now11>" & vbCrLf)
                Catch ex As Exception
                    Throw ex
                    'xmlsb.Append(vbCrLf & "<error>" & vbCrLf)
                    'xmlsb.Append(ex)
                    'xmlsb.Append(vbCrLf & "</error>" & vbCrLf)
                    'xmlsb.Append(vbCrLf & "<sql>" & vbCrLf)
                    'xmlsb.Append(sql)
                    'xmlsb.Append(vbCrLf & "</sql>" & vbCrLf)
                Finally
                    dbconn1.Close()
                End Try

                Dim oxml As XmlDocument = New XmlDocument
                Dim xRss As XmlNode

                For Each xRss In xDataSet.SelectNodes("RSS")

                    rssurl = CommFunc.nullText(xRss.SelectSingleNode("URL"))

                    xmlsb.Append(vbCrLf & "<rssurl>" & vbCrLf)
                    xmlsb.Append(rssurl)
                    xmlsb.Append(vbCrLf & "</rssurl>" & vbCrLf)

                    oxml.Load(rssurl)

                    Dim RssTop As Integer = CInt(CommFunc.nullText(xRss.SelectSingleNode("RssTop")))
                    Dim xnewWindow As String = CommFunc.nullText(xRss.SelectSingleNode("newWindow"))
                    If xnewWindow = "" Then xnewWindow = "N"
                    Dim xRssCount As Integer = 0
                    Dim xItem As XmlNode
                    For Each xItem In oxml.SelectNodes("rss/channel/item")
                        If xRssCount < RssTop Then
                            xRssCount = xRssCount + 1
                            
                            xmlsb.Append("<Article iCuItem=""0"" newWindow=""" & xnewWindow & """>")
                            'xmlsb.Append("<category>" & CommFunc.nullText(xItem.SelectSingleNode("category")) & "</category>")
                            xmlsb.Append("<Caption>" & CommFunc.nullText(xItem.SelectSingleNode("title")) & "</Caption>")
                            If CommFunc.nullText(xItem.SelectSingleNode("description")) <> "" Then
                                xmlsb.Append("<Abstract><![CDATA[" & CommFunc.nullText(xItem.SelectSingleNode("description")) & "]]></Abstract>")
                            End If
                            If CommFunc.nullText(xItem.SelectSingleNode("Content")) <> "" Then
                                xmlsb.Append("<Content><![CDATA[" & CommFunc.nullText(xItem.SelectSingleNode("Content")) & "]]></Content>")
                            End If

                            'Dim pubDate As String = Mid(CommFunc.nullText(xItem.SelectSingleNode("pubDate")), 5, 12)
                            Dim pubDate As String = CommFunc.nullText(xItem.SelectSingleNode("pubDate"))
                            Dim MyShortDate As Date = CDate(pubDate)

                            xmlsb.Append("<PostDate>" & MyShortDate & "</PostDate>")
                            'xmlsb.Append("<DeptName>" & nullText(xItem.SelectSingleNode("author")) & "</DeptName>")
                            xmlsb.Append("<xURL><![CDATA[" & CommFunc.deAmp(CommFunc.nullText(xItem.SelectSingleNode("link"))) & "]]></xURL>")
                            'xmlsb.Append("<xImgFile>" & deAmp(nullText(xItem.SelectSingleNode("img"))) & "</xImgFile>")
                            xmlsb.Append("</Article>")
                        End If
                    Next
                Next
                xmlsb.Append("</" & CommFunc.nullText(xDataSet.SelectSingleNode("DataLable")) & ">")
            End If

        Catch ex As Exception
            Throw ex
            'xmlsb.Append(vbCrLf & "<error>" & vbCrLf)
            'xmlsb.Append(ex)
            'xmlsb.Append(vbCrLf & "</error>" & vbCrLf)
            'xmlsb.Append(vbCrLf & "<rssurl>" & vbCrLf)
            'xmlsb.Append(rssurl)
            'xmlsb.Append(vbCrLf & "</rssurl>" & vbCrLf)
        Finally
            dbconn.Close()
        End Try

        Return xmlsb.ToString()

    End Function
End Class
