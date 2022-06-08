Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class Menus
    Shared Function x1Menus(ByVal refModel As XmlNode, ByVal myTreeNode As Integer, ByVal upParent As Integer, ByVal topParent As Integer, ByVal mp As String) As String
        'Dim retString As String = ""
        Dim xmlsb As StringBuilder = New StringBuilder
        Dim MenuTree As String
        Dim xRootID As String
        Dim sqlstr_MenuTree As String
        '----MenuTree
        MenuTree = ""
        xRootID = CommFunc.nullText(refModel.SelectSingleNode("MenuTree"))
        If xRootID <> "" Then
            sqlstr_MenuTree = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
                    & " FROM CatTreeNode AS a " _
                    & " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
                    & " WHERE a.inUse='Y' AND a.DataParent=" & upParent _
                    & " AND a.CtRootID = @xRootID" _
                    & " Order by a.CatShowOrder"
            xmlsb.Append(x1Menus_Detail(MenuTree, sqlstr_MenuTree, myTreeNode, xRootID, mp, True))
        End If
        '----MenuTree1
        MenuTree = "1"
        xRootID = CommFunc.nullText(refModel.SelectSingleNode("MenuTree1"))
        If xRootID <> "" Then
            sqlstr_MenuTree = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
                & " FROM CatTreeNode AS a " _
                & " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
                & " WHERE a.inUse='Y' AND a.DataLevel=1 AND a.CtRootID = @xRootID" _
                & " Order by a.CatShowOrder"
            xmlsb.Append(x1Menus_Detail(MenuTree, sqlstr_MenuTree, myTreeNode, xRootID, mp, False))
        End If
        '----MenuTree2
        MenuTree = "2"
        xRootID = CommFunc.nullText(refModel.SelectSingleNode("MenuTree2"))
        If xRootID <> "" Then
            sqlstr_MenuTree = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
                & " FROM CatTreeNode AS a " _
                & " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
                & " WHERE a.inUse='Y' AND a.DataLevel=1 AND a.CtRootID = @xRootID" _
                & " Order by a.CatShowOrder"
            xmlsb.Append(x1Menus_Detail(MenuTree, sqlstr_MenuTree, myTreeNode, xRootID, mp, True))
        End If

        'Return retString
        Return xmlsb.ToString
    End Function

    Shared Function x1Menus_Detail(ByVal MenuTree As String, ByVal sqlstr_MenuTree As String, ByVal myTreeNode As Integer, ByVal xRootID As String, ByVal mp As String, ByVal alwaysShowChild As Boolean) As String
        Dim connstringName As String = WebConfigurationManager.AppSettings("connstringName")
        Dim connstringString = WebConfigurationManager.ConnectionStrings(connstringName).ConnectionString
        'Dim retString As String = ""
        Dim xmlsb As StringBuilder = New StringBuilder
        Dim xURL As String = ""
        Dim SQLCom As String = sqlstr_MenuTree
        Dim myConnection As New SqlConnection(connstringString)
        Dim myCommand As New SqlCommand(SQLCom, myConnection)
        myCommand.Parameters.AddWithValue("@xRootID", xRootID)
        myConnection.Open()
        Dim myReader As SqlDataReader
        myReader = myCommand.ExecuteReader()

        xmlsb.Append("<MenuBar" + MenuTree + " myTreeNode=""" & myTreeNode & """>")
        If myReader.HasRows Then
            Dim tabindex As Integer = 0
            Dim divindex As Integer = 0
            Dim styletop As Integer = 207
            While myReader.Read() = True
                tabindex = tabindex + 1
                divindex = divindex + 1
                styletop = styletop + 28

                Dim ctUnitKind As String = CommFunc.getValue(myReader("CtUnitKind"))

                If ctUnitKind = "K" Then
                    xURL = "kmDoit.asp?" & myReader("redirectURL")
                ElseIf myReader("CtNodeKind") = "C" Then  '-- Folder
                    xURL = "np.asp?ctNode=" & myReader("CtNodeID") & "&mp=" & mp
                Else
                    Dim redirectURL As String = CommFunc.getValue(myReader("redirectURL"))

                    If redirectURL <> "" Then
                        If redirectURL.IndexOf("?") < 0 Then
                            redirectURL = redirectURL & "?"
                        End If
                        If redirectURL.IndexOf("mp=") < 0 Then
                            redirectURL = redirectURL & "&mp=" & mp
                        End If
                        xURL = redirectURL
                    ElseIf ctUnitKind = "2" Then
                        xURL = "lp.asp?ctNode=" & myReader("CtNodeID") & "&CtUnit=" & myReader("CtUnitID") & "&BaseDSD=" & myReader("iBaseDSD") & "&mp=" & mp
                    ElseIf IsNumeric(myReader("iBaseDSD")) Then
                        xURL = "np.asp?ctNode=" & myReader("CtNodeID") & "&CtUnit=" & myReader("CtUnitID") & "&BaseDSD=" & myReader("iBaseDSD") & "&mp=" & mp
                    End If
                End If

                xmlsb.Append("<MenuCat newWindow=""" & myReader("newWindow") & """ xNode=""" & myReader("CtNodeID") & """>")
                xmlsb.Append("<Caption>" & CommFunc.deAmp(myReader("CatName")) & "</Caption>")
                xmlsb.Append("<redirectURL>" & CommFunc.deAmp(xURL) & "</redirectURL>")
                xmlsb.Append("<tabindex>" & tabindex & "</tabindex>")
                xmlsb.Append("<divindex>div" & divindex & "</divindex>")
                xmlsb.Append("<Mover>MM_showHideLayers('div" & divindex & "','','show')</Mover>")
                xmlsb.Append("<Mout>MM_showHideLayers('div" & divindex & "','','hide')</Mout>")
                xmlsb.Append("<styletop>" & styletop & "</styletop>")

                If myReader("CtNodeID") = CInt(myTreeNode) Or AlwaysShowChild = True Then
                    SQLCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
                   & " FROM CatTreeNode AS a " _
                   & " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
                   & " WHERE a.inUse='Y' AND a.DataParent=@myTreeNode AND a.CtRootID = @xRootID" _
                   & " Order by a.CatShowOrder"

                    Dim myConnection1 As New SqlConnection(connstringString)
                    Dim myCommand1 As New SqlCommand(SQLCom, myConnection1)
                    myCommand1.Parameters.AddWithValue("@myTreeNode", myTreeNode)
                    myCommand1.Parameters.AddWithValue("@xRootID", xRootID)
                    myConnection1.Open()
                    Dim myReader1 As SqlDataReader
                    myReader1 = myCommand1.ExecuteReader()

                    If myReader1.HasRows Then
                        While myReader1.Read() = True
                            tabindex = tabindex + 1
                            If myReader("CtNodeID") = CInt(myTreeNode) Then
                                styletop = styletop + 26
                            End If
                            Dim ctUnitKind1 As String = CommFunc.getValue(myReader1("CtUnitKind"))
                            Dim redirectURL As String = CommFunc.getValue(myReader1("redirectURL"))

                            If ctUnitKind1 = "K" Then
                                xURL = "kmDoit.asp?" & redirectURL
                            ElseIf myReader1("CtNodeKind") = "C" Then  '-- Folder
                                xURL = "np.asp?ctNode=" & myReader1("CtNodeID") & "&mp=" & mp
                            Else
                                If redirectURL <> "" Then
                                    xURL = redirectURL
                                ElseIf ctUnitKind1 = "2" Then
                                    xURL = "lp.asp?ctNode=" & myReader1("CtNodeID") & "&CtUnit=" & myReader1("CtUnitID") & "&BaseDSD=" & myReader1("iBaseDSD") & "&mp=" & mp
                                ElseIf IsNumeric(myReader1("iBaseDSD")) Then
                                    xURL = "np.asp?ctNode=" & myReader1("CtNodeID") & "&CtUnit=" & myReader1("CtUnitID") & "&BaseDSD=" & myReader1("iBaseDSD") & "&mp=" & mp
                                End If
                            End If

                            xmlsb.Append("<MenuItem newWindow=""" & myReader1("newWindow") & """ xNode=""" & myReader1("CtNodeID") & """>")
                            xmlsb.Append("<Caption>" & CommFunc.deAmp(myReader1("CatName")) & "</Caption>")
                            xmlsb.Append("<redirectURL>" & CommFunc.deAmp(xURL) & "</redirectURL>")
                            xmlsb.Append("<tabindex>" & tabindex & "</tabindex>")
                            xmlsb.Append("</MenuItem>")

                        End While
                    End If
                    myReader1.Close()
                    myConnection1.Close()
                End If

                xmlsb.Append("</MenuCat>")

            End While
        End If

        xmlsb.Append("</MenuBar" + MenuTree + ">")

        myReader.Close()
        myConnection.Close()
        myConnection = Nothing
        'Return retString
        Return xmlsb.ToString
    End Function
End Class
