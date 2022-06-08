Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class KnowledgeUtility

    Dim iCuitem As String

    Public Sub New(ByVal item As String)
        iCuitem = item
    End Sub

    Public Function GetDiscussCount() As String

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection = New SqlConnection(ConnString)
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim totalDiscuss As String = ""


        Try


            sqlString = "SELECT COUNT(*) FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
            sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
            sqlString &= "AND (KnowledgeForum.Status = 'N') "

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", iCuitem)
            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))

            
            totalDiscuss = myCommand.ExecuteScalar

            '---debug---
            'System.Web.HttpContext.Current.Response.write(sqlString & "<br/>")
            'System.Web.HttpContext.Current.Response.write("iCuitem:" & iCuitem & "<br/>")
            'System.Web.HttpContext.Current.Response.write("ictunit:" & WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId") & "<br/>")
            'System.Web.HttpContext.Current.Response.write("SiteId:" & WebConfigurationManager.AppSettings("SiteId") & "<br/>")
            'System.Web.HttpContext.Current.Response.write(totalDiscuss & "<br/>")

            Return totalDiscuss

            myCommand.Dispose()
        Catch ex As Exception
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Function

End Class
