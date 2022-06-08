
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class PediaUtility
    Public Shared Function ReplacePedia(ByVal strTemp As String) As Array
        Dim pediaTagInfo(1) As String   '0擺放要取代的文章內容 1擺放隱藏的table
        Dim pediaTableHtml As String
        Dim aid As String = System.Configuration.ConfigurationSettings.AppSettings("PediaCtUnitId")
        Dim sqlString As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim cnn As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim myTable As DataTable
        Dim myDataRow As DataRow
        Dim total As Integer = 0
        Dim oCache As System.Web.Caching.Cache
        oCache = System.Web.HttpContext.Current.Cache
        If oCache("Pedia") Is Nothing Then
            cnn = New SqlConnection(ConnString)
            Try
                sqlString = "SELECT CuDTGeneric.iCUItem,CuDTGeneric.sTitle FROM CuDTGeneric "
                sqlString &= "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
                sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y')"

                cnn.Open()
                myCommand = New SqlCommand(sqlString, cnn)
                myCommand.Parameters.AddWithValue("@ictunit", aid)
                myReader = myCommand.ExecuteReader()

                myTable = New DataTable()
                myTable.Load(myReader)
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()
                cnn.Close()
            Catch ex As Exception
            End Try
            oCache.Insert("Pedia", myTable, _
                                Nothing, _
                                System.Web.Caching.Cache.NoAbsoluteExpiration, New TimeSpan(0, 5, 0))
        Else
            myTable = oCache("Pedia")
        End If
        total = myTable.Rows().Count()
        '比對文章內有無農業詞彙
        While total <> 0
            total = total - 1
            myDataRow = myTable.Rows.Item(total)
            If strTemp.IndexOf(myDataRow.Item("sTitle")) <> -1 Then
                '取代要取代的辭彙給一個id用來呼叫 pedia.ja 移動table與顯示隱藏用
                'strTemp = Replace(strTemp, myDataRow.Item("sTitle"), "<a id=""ine_" & myDataRow("iCUItem") & """ onclick='controlDiv(""ine_" & myDataRow("iCUItem") & """,this);' class=""pediaword"" style=""color:purple;""><span style='background:yellow'>" & myDataRow.Item("sTitle") + "</span></a>", 1, 1)
                'Modify by Max 2011/06/03
                '此為新版的詞彙程式,對應到subject/ShowPhrase.asp這一隻程式
                strTemp = Replace(strTemp, myDataRow.Item("sTitle"), "<a id='ine_" & Guid.NewGuid().tostring() & "' class='jTip jTip1' name='../subject/ShowPhrase.asp?width=520&iCUItem=" & myDataRow("iCUItem") & "'><span>" & myDataRow.Item("sTitle") & "</span></a>", 1, 1)
				'End
				'將對應的辭彙增加一個table 顯示其內容
                pediaTableHtml &= WritePediaTag(myDataRow("iCUItem"))
            End If
        End While

        pediaTagInfo(1) = pediaTableHtml
        pediaTagInfo(0) = strTemp
        Return pediaTagInfo
    End Function

    '寫出要顯示的內容 sb1是詞彙資訊 sb2是補充內容
    Public Shared Function WritePediaTag(ByVal tagId As String) As String
        Dim sb1 As String = GetPediaTable(tagId)
        Dim sb2 As String = ""
        If Not sb1.Contains("名詞釋義") Then
            sb2 = GetPediaAdded(tagId)
        End If
        Dim htmtemp As String = ""
        htmtemp &= "<Table id='inc_ine_" & tagId & "' class='aas'><tbody>"
        htmtemp &= "<tr><td align=""right""><div class=""aasclose"" onclick='closeDiv(""ine_" & tagId & """);'>X</div></td></tr>"
        htmtemp &= "<tr><td>" & sb1 & "</td></tr><tr><td>" & sb2 & "</td></tr><tr><td align='right'><a target='_blank' href='/Pedia/PediaContent.aspx?AId=" & tagId & "'>more..</a></td></tr></tbody></table>"
        Return htmtemp
    End Function

    '辭彙內容區塊
    Public Shared Function GetPediaTable(ByVal pediaId As String) As String
        Dim sqlString As String = ""
        Dim tableHtm As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim cnn As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim total As Integer = 0
        cnn = New SqlConnection(ConnString)
        Try
            sqlString = "  SELECT CuDTGeneric.sTitle, CuDTGeneric.xBody, CuDTGeneric.xKeyword, CuDTGeneric.vGroup, Pedia.engTitle, "
            sqlString &= " Pedia.formalName, Pedia.localName FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
            sqlString &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem)"

            cnn.Open()
            myCommand = New SqlCommand(sqlString, cnn)
            myCommand.Parameters.AddWithValue("@icuitem", pediaId)
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                tableHtm &= "<table class=""explained"">"
                tableHtm &= "<tr><th width=""40%"">詞目</th><td>" & myReader("sTitle") & "</td></tr>"
                If Not String.IsNullOrEmpty(myReader("engTitle")) Then
                    tableHtm &= "<tr><th>英文詞目</th><td>" & IIf(myReader("engTitle") Is DBNull.Value Or myReader("engTitle") = "", "&nbsp;", myReader("engTitle")) & "</td></tr>"
                End If
                'tableHtm &= "<tr><th>學名(中/英)</th><td>" & IIf(myReader("formalName") Is DBNull.Value Or myReader("formalName") = "", "&nbsp;", myReader("formalName")) & "</td></tr>"
                'tableHtm &= "<tr><th>俗名(中/英)</th><td>" & IIf(myReader("localName") Is DBNull.Value Or myReader("localName") = "", "&nbsp;", myReader("localName")) & "</td></tr>"
                If Not String.IsNullOrEmpty(myReader("xBody")) Then

                    tableHtm &= "<tr><th>名詞釋義</th><td>" & IIf(myReader("xBody") Is DBNull.Value Or myReader("xBody") = "", "&nbsp;", myReader("xBody").ToString.Replace(vbCrLf, "<br />")) & "</td></tr>"
                End If
                If Not String.IsNullOrEmpty(myReader("xKeyword")) Then
                    '    tableHtm &= "<tr><th>相關詞</th><td>&nbsp;</td></tr>"
                    'Else
                    Dim items = myReader("xKeyword").ToString.Split(";")
                    tableHtm &= "<tr><th>相關詞</th><td>"
                    For Each item As String In items
                        If item IsNot Nothing And item <> "" Then
                            tableHtm &= item & ","
                        End If
                    Next
                    tableHtm &= "</td></tr>"
                End If
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

    '補充解釋區塊
    Public Shared Function GetPediaAdded(ByVal addedId As String) As String
        Dim tableHtm As String = ""
        Dim sqlString As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim cnn As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim index As Integer = 1
        cnn = New SqlConnection(ConnString)
        Try
            sqlString = "  SELECT TOP (3) CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xBody, Pedia.parentIcuItem, Member.realname, Member.nickname, "
            sqlString &= " Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
            sqlString &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.parentIcuitem = @icuitem) ORDER BY Pedia.commendTime DESC"

            cnn.Open()
            myCommand = New SqlCommand(sqlString, cnn)
            myCommand.Parameters.AddWithValue("@icuitem", addedId)
            myReader = myCommand.ExecuteReader

            tableHtm &= "<table class=""pediaadded"">"
            tableHtm &= "<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"">補充解釋(摘要)</th><th scope=""col"" width=""10%"">發表者</th><th scope=""col"" width=""10%"">發佈日期</th></tr>"
            While myReader.Read
                tableHtm &= "<tr><td>" & index & "</td>"
                tableHtm &= "<td><a href=""/Pedia/PediaExplainContent.aspx?AId=" & myReader("iCUItem") & "&PAId=" & myReader("parentIcuItem") & """ target=""_blank"">"
                If myReader("xBody").ToString.Length > 70 Then
                    tableHtm &= myReader("xBody").ToString.Substring(0, 70)
                Else
                    tableHtm &= myReader("xBody").ToString
                End If
                tableHtm &= "</a></td>"
                If myReader("nickname") IsNot DBNull.Value Then
                    If myReader("nickname") <> "" Then
                        tableHtm &= "<td>" & myReader("nickname").ToString & "</td>"
                    Else
                        tableHtm &= "<td>" & myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>"
                    End If
                ElseIf myReader("realname") IsNot DBNull.Value Then
                    tableHtm &= "<td>" & myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>"
                Else
                    tableHtm &= "<td>&nbsp;</td>"
                End If
                tableHtm &= "<td>" & Date.Parse(myReader("commendTime")).ToShortDateString & "</td>"
                tableHtm &= "</tr>"
                index += 1
            End While
            tableHtm &= "</table>"
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
        End Try
        Return tableHtm
    End Function
End Class