Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Public Class KPI_Login

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim myTable As DataTable
    Dim myDataRow As DataRow

    Dim _memberId As String
    Dim _ctNode As String
    Dim _siteId As String
    Dim _pageType As String
    Dim _memberGrade As Integer
    Dim _loginGrade As Integer
    Dim _loginCode As String
    Dim _loginCode2 As String

    Public Sub New(ByVal siteId As String, ByVal memberId As String, ByVal ctNode As String, ByVal pageType As String)

        _siteId = siteId
        _memberId = memberId
        _ctNode = ctNode
        _pageType = pageType
        _memberGrade = 0
        _loginGrade = 0
        _loginCode = "st_3"

        '會員登入  st_311
        'IQ大挑戰登入  st_312
        '虛擬養殖登入   st_313


        If _siteId = "0" Then
            _loginCode2 = "st_311"   '---會員登入---
        ElseIf _siteId = "1" Then
            _loginCode2 = "st_312"   '---大挑戰登入---
        ElseIf _siteId = "2" Then
            _loginCode2 = "st_313"   '---虛擬養殖登入---
      
        Else
            Exit Sub
        End If

    End Sub

    Public Sub HandleLogin()

        '---檢查瀏覽行為是否公開---
        If Not CheckLoginIsPublic() Then

            Exit Sub
        End If

        '---檢查傳入的memberId是否有登入當時的資料---
        If Not CheckMemberLoginToday() Then
            '---此member沒有今天的記錄---插入一筆新的---

            If Not InsertMemberDefault() Then
                Exit Sub
            End If
        Else
            '---此member有今天的資料---檢查今日分數是否超過上限---
            If Not CheckGradeLimit() Then


                Exit Sub
            End If
        End If

        '---檢查傳入的ctnode是否有值---
        If _ctNode = "" Then
            '---利用siteid來抓值---
            If Not GetLoginGradeBySiteId() Then
                Exit Sub
            End If
        Else
            '---利用ctnode來抓值---
            'If Not GetLoginGradeByCtNode() Then
            '    '---若此ctnode沒有特殊設定,利用siteid抓值---
            '    If Not GetLoginGradeBySiteId() Then
            '        Exit Sub
            '    End If
            'End If
        End If

        '---利用得到的給點,存入table中---
        If Not SaveGradeIntoDB() Then
            Exit Sub
        End If

    End Sub

    ''' <summary>檢查瀏覽行為是否公開</summary>
    ''' <returns>公開 return true, 關閉 return false</returns>
    ''' <remarks></remarks>
    Function CheckLoginIsPublic() As Boolean

        Dim returnStr As String = ""
        returnStr = Trim(GetID.Get_OneField_ReturnIdStr("kpi_set_ind", "Rank0", "st_3", "Rank2"))

        If returnStr = "Y" Then
            Return True
        Else
            Return False
        End If


    End Function

    Function CheckMemberLoginToday() As Boolean

        '今天有登入  RETURN  TRUE

        Dim returnStr As String = ""


        Dim MyCn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Dim sql As String = ""
        Dim IDStr As String = ""
        Dim reader As SqlDataReader
        Dim FLAG As Boolean = False



        Dim previousConnectionState As ConnectionState = MyCn.State
        Dim Command As SqlCommand = New SqlCommand

        Try
            Command.CommandType = CommandType.Text

            sql = "SELECT     memberId, loginDate  FROM MemberGradeLogin "
            sql += " where  memberId=@memberId and CONVERT(varchar, loginDate, 111) =CONVERT(varchar, getdate(), 111)"

            Command.CommandText = sql
            Command.Parameters.Add(New SqlParameter("@memberId", _memberId))


            MyCn.Open()
            Command.Connection = MyCn
            reader = Command.ExecuteReader

            If reader.Read Then
                FLAG = True
            End If

            Return FLAG

        Catch ex As Exception

            Return False

        Finally
            If MyCn.State = ConnectionState.Open Then
                MyCn.Close()
            End If
        End Try



    End Function

    Function InsertMemberDefault() As Boolean


        Dim MyCn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Dim previousConnectionState As ConnectionState = MyCn.State
        Dim Command As SqlCommand = New SqlCommand
        Dim TempDate As Date = Today


        Try
            If MyCn.State = ConnectionState.Closed Then
                MyCn.Open()
            End If

            Dim sql As String = ""
            sql = "Insert  MemberGradeLogin (memberId, loginDate)"
            sql += " Values (@memberId,  getdate())"


            Command.Parameters.Add(New SqlParameter("@memberId", _memberId))


            Command.Connection = MyCn
            Command.CommandText = sql
            Command.ExecuteNonQuery()
            Return True


        Catch ex As Exception
            Return False
        Finally
            If previousConnectionState = ConnectionState.Closed Then
                MyCn.Close()
            End If
        End Try

    End Function

    ''' <summary>檢查瀏覽分數總合是否超過當日上限</summary>
    ''' <returns>超過上限return false, 沒超過上限return true</returns>
    ''' <remarks></remarks>
    Function CheckGradeLimit() As Boolean

        '會員登入  st_311
        'IQ大挑戰登入  st_312
        '虛擬養殖登入   st_313

        Dim LimitScore As Integer = GetID.Get_TwoField("kpi_set_score", "Rank0", "st_3", "Rank0_2", _loginCode2, "Rank0_1")

     

        Dim TotalScore As Integer = 0

        Dim MyCn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Dim sql As String = ""
        Dim IDStr As String = ""
        Dim reader As SqlDataReader
        Dim FLAG As Boolean = False

        '時間放在第二個欄位


        Dim previousConnectionState As ConnectionState = MyCn.State
        Dim Command As SqlCommand = New SqlCommand

        Try
            Command.CommandType = CommandType.Text

            sql = "SELECT    loginInterGrade as Total "
            sql += " from   MemberGradeLogin where  memberId=@memberId and CONVERT(varchar, LoginDate, 111) =CONVERT(varchar, getdate(), 111)"

            Command.CommandText = sql
            Command.Parameters.Add(New SqlParameter("@memberId", _memberId))
            MyCn.Open()
            Command.Connection = MyCn
            reader = Command.ExecuteReader

            If reader.Read Then
                If Not (reader.Item("Total") Is System.DBNull.Value) Then
                    If Trim(reader.Item("Total")) <> "" Then
                        TotalScore = CInt(reader.Item("Total"))
                    End If

                End If
            End If


            If TotalScore > LimitScore Then
                Return False
            Else
                Return True

            End If

        Catch ex As Exception

            Return False

        Finally
            If MyCn.State = ConnectionState.Open Then
                MyCn.Close()
            End If
        End Try




    End Function

    ''' <summary>利用ctnode的值來找出特殊的分數</summary>  
    ''' <returns>若有找到ctnode相對應的分數 return true, 若沒有 return false</returns>
    ''' <remarks></remarks>
    Function GetLoginGradeByCtNode() As Boolean

    End Function

    Function GetLoginGradeBySiteId() As Boolean
        Try

            _loginGrade = GetID.Get_OneField_ReturnIdStr("kpi_set_score", "Rank0_2", _loginCode2, "Rank0_1")

        Catch
            Return False

        End Try

        Return True

    End Function

    Function SaveGradeIntoDB() As Boolean

        Dim MyCn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Dim previousConnectionState As ConnectionState = MyCn.State
        Dim Command As SqlCommand = New SqlCommand



        '會員登入  st_311
        'IQ大挑戰登入  st_312
        '虛擬養殖登入   st_313

        Try
            If MyCn.State = ConnectionState.Closed Then
                MyCn.Open()
            End If

            Dim sql As String = ""


            sql = "Update  MemberGradeLogin  set "


            If _loginCode2 = "st_311" Then

                sql += " loginInterCount=(loginInterCount+1), loginInterGrade=(loginInterGrade+@loginInterGrade), loginInterDate=getdate()"
                Command.Parameters.Add(New SqlParameter("@loginInterGrade", _loginGrade))
              

            ElseIf _loginCode2 = "st_312" Then


                sql += " loginIQCount=(loginIQCount+1), loginIQGrade=(loginIQGrade+@loginIQGrade), loginIQDate=getdate()"
                Command.Parameters.Add(New SqlParameter("@loginIQCount", _loginGrade))

            ElseIf _loginCode2 = "st_313" Then
                sql += " loginRaiseCount=(loginIQCount+1), loginRaiseGrade=(loginRaiseGrade+@loginRaiseGrade), loginRaiseDate=getdate()"
                Command.Parameters.Add(New SqlParameter("@loginRaiseGrade", _loginGrade))


            
            End If


            sql += " where  memberId=@memberId and CONVERT(varchar, LoginDate, 111) =CONVERT(varchar, getdate(), 111)"


            Command.Parameters.Add(New SqlParameter("@memberId", _memberId))


            Command.Connection = MyCn
            Command.CommandText = sql
            Command.ExecuteNonQuery()

            Return True



        Catch ex As Exception
            Return False

        Finally
            If previousConnectionState = ConnectionState.Closed Then
                MyCn.Close()
            End If
        End Try

    End Function

End Class
