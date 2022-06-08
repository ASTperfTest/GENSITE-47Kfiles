
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class PediaUtility
    Public Shared Function ReplacePedia(ByVal strTemp As String) As Array
        Dim pediaTagInfo(1) As String   '0�\��n���N���峹���e 1�\�����ê�table
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
        '���峹�����L�A�~���J
        While total <> 0
            total = total - 1
            myDataRow = myTable.Rows.Item(total)
            If strTemp.IndexOf(myDataRow.Item("sTitle")) <> -1 Then
                '���N�n���N����J���@��id�ΨөI�s pedia.ja ����table�P������å�
                'strTemp = Replace(strTemp, myDataRow.Item("sTitle"), "<a id=""ine_" & myDataRow("iCUItem") & """ onclick='controlDiv(""ine_" & myDataRow("iCUItem") & """,this);' class=""pediaword"" style=""color:purple;""><span style='background:yellow'>" & myDataRow.Item("sTitle") + "</span></a>", 1, 1)
                'Modify by Max 2011/06/03
                '�����s�������J�{��,������subject/ShowPhrase.asp�o�@���{��
                strTemp = Replace(strTemp, myDataRow.Item("sTitle"), "<a id='ine_" & Guid.NewGuid().tostring() & "' class='jTip jTip1' name='../subject/ShowPhrase.asp?width=520&iCUItem=" & myDataRow("iCUItem") & "'><span>" & myDataRow.Item("sTitle") & "</span></a>", 1, 1)
				'End
				'�N��������J�W�[�@��table ��ܨ䤺�e
                pediaTableHtml &= WritePediaTag(myDataRow("iCUItem"))
            End If
        End While

        pediaTagInfo(1) = pediaTableHtml
        pediaTagInfo(0) = strTemp
        Return pediaTagInfo
    End Function

    '�g�X�n��ܪ����e sb1�O���J��T sb2�O�ɥR���e
    Public Shared Function WritePediaTag(ByVal tagId As String) As String
        Dim sb1 As String = GetPediaTable(tagId)
        Dim sb2 As String = ""
        If Not sb1.Contains("�W�����q") Then
            sb2 = GetPediaAdded(tagId)
        End If
        Dim htmtemp As String = ""
        htmtemp &= "<Table id='inc_ine_" & tagId & "' class='aas'><tbody>"
        htmtemp &= "<tr><td align=""right""><div class=""aasclose"" onclick='closeDiv(""ine_" & tagId & """);'>X</div></td></tr>"
        htmtemp &= "<tr><td>" & sb1 & "</td></tr><tr><td>" & sb2 & "</td></tr><tr><td align='right'><a target='_blank' href='/Pedia/PediaContent.aspx?AId=" & tagId & "'>more..</a></td></tr></tbody></table>"
        Return htmtemp
    End Function

    '��J���e�϶�
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
                tableHtm &= "<tr><th width=""40%"">����</th><td>" & myReader("sTitle") & "</td></tr>"
                If Not String.IsNullOrEmpty(myReader("engTitle")) Then
                    tableHtm &= "<tr><th>�^�����</th><td>" & IIf(myReader("engTitle") Is DBNull.Value Or myReader("engTitle") = "", "&nbsp;", myReader("engTitle")) & "</td></tr>"
                End If
                'tableHtm &= "<tr><th>�ǦW(��/�^)</th><td>" & IIf(myReader("formalName") Is DBNull.Value Or myReader("formalName") = "", "&nbsp;", myReader("formalName")) & "</td></tr>"
                'tableHtm &= "<tr><th>�U�W(��/�^)</th><td>" & IIf(myReader("localName") Is DBNull.Value Or myReader("localName") = "", "&nbsp;", myReader("localName")) & "</td></tr>"
                If Not String.IsNullOrEmpty(myReader("xBody")) Then

                    tableHtm &= "<tr><th>�W�����q</th><td>" & IIf(myReader("xBody") Is DBNull.Value Or myReader("xBody") = "", "&nbsp;", myReader("xBody").ToString.Replace(vbCrLf, "<br />")) & "</td></tr>"
                End If
                If Not String.IsNullOrEmpty(myReader("xKeyword")) Then
                    '    tableHtm &= "<tr><th>������</th><td>&nbsp;</td></tr>"
                    'Else
                    Dim items = myReader("xKeyword").ToString.Split(";")
                    tableHtm &= "<tr><th>������</th><td>"
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

    '�ɥR�����϶�
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
            tableHtm &= "<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"">�ɥR����(�K�n)</th><th scope=""col"" width=""10%"">�o���</th><th scope=""col"" width=""10%"">�o�G���</th></tr>"
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
                        tableHtm &= "<td>" & myReader("realname").ToString.Substring(0, 1) & "��" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>"
                    End If
                ElseIf myReader("realname") IsNot DBNull.Value Then
                    tableHtm &= "<td>" & myReader("realname").ToString.Substring(0, 1) & "��" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>"
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