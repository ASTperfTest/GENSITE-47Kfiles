Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class GetID

  Public Overloads Shared Function Get_OneField_ReturnIdStr(ByVal TableName As String, ByVal fieldName As String, _
                                                            ByVal fieldvalue As String, ByVal SearchfieldName As String) As String

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Dim IDStr As String = ""

    myConnection = New SqlConnection(ConnString)

    Try
      sql = " SELECT " & SearchfieldName & " FROM " & TableName & " WHERE " & fieldName & " = @" & fieldName
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@" & fieldName, fieldvalue)
      myReader = myCommand.ExecuteReader

      While myReader.Read
        IDStr &= CStr(myReader.Item(0)) & ","
      End While

      If IDStr <> "" Then
        IDStr = Left(IDStr, Len(IDStr) - 1)
      End If

      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return IDStr

  End Function

  Public Overloads Shared Function Get_TwoField(ByVal TableName As String, ByVal fieldName1 As String, _
                                                ByVal fieldvalue1 As String, ByVal fieldName2 As String, _
                                                ByVal fieldvalue2 As String, ByVal SearchfieldName As String) As Long

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim IDStr As String = ""

    myConnection = New SqlConnection(ConnString)
    Try

      sql = " SELECT " & SearchfieldName & " FROM " & TableName & " WHERE " & fieldName1 & " = @" & fieldName1
      sql &= " AND " & fieldName2 & " = @" & fieldName2

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@" & fieldName1, fieldvalue1)
      myCommand.Parameters.AddWithValue("@" & fieldName2, fieldvalue2)

      myReader = myCommand.ExecuteReader

      While myReader.Read
        IDStr &= CStr(myReader.Item(0)) & ","
      End While

      If IDStr <> "" Then
        IDStr = Left(IDStr, Len(IDStr) - 1)
      End If

      If Not myReader.IsClosed Then
        myReader.Close()
      End If

      myCommand.Dispose()

    Catch ex As Exception
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return IDStr

  End Function

End Class

