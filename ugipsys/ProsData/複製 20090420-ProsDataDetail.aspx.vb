Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class ProsData_ProsDataDetail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then


            Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
            Dim conn As SqlConnection
            Dim comm As SqlCommand
            Dim reader As SqlDataReader
            Dim sql As String

            '---check this account is used or not---
            conn = New SqlConnection(ConnString)
            Try
                conn.Open()
                sql = "SELECT * FROM Member WHERE account = @account"
                comm = New SqlCommand(sql, conn)
                comm.Parameters.AddWithValue("@account", Request.QueryString("account"))
                reader = comm.ExecuteReader
                If reader.Read Then

                    accounttext.Text = reader("account").ToString
                    realnametext.Text = reader("realname").ToString
                    nicknametext.Text = reader("nickname").ToString
                    password1text.Text = ""
                    password2text.Text = ""
                    idntext.Text = reader("id").ToString
                    member_orgtext.Text = reader("member_org").ToString
                    com_teltext.Text = reader("com_tel").ToString
                    com_exttext.Text = reader("com_ext").ToString
                    ptitletext.Text = reader("ptitle").ToString
                    If reader("birthday") IsNot DBNull.Value Or reader("birthday") <> "" Then
                        birthYeartext.Text = reader("birthday").ToString.Substring(0, 4)
                        birthMonthtext.Items.FindByValue(reader("birthday").ToString.Substring(4, 2)).Selected = True
                        birthdaytext.Items.FindByValue(reader("birthday").ToString.Substring(6, 2)).Selected = True
                    Else
                        birthYeartext.Text = ""
                    End If

                    If reader("sex").ToString = "1" Then
                        maletext.Checked = True
                    ElseIf reader("sex").ToString = "0" Then
                        femaletext.Checked = True
                    End If
                    homeaddrtext.Text = reader("homeaddr").ToString
                    ziptext.Text = reader("zip").ToString
                    phonetext.Text = reader("phone").ToString
                    home_exttext.Text = reader("home_ext").ToString
                    mobiletext.Text = reader("mobile").ToString
                    faxtext.Text = reader("fax").ToString
                    emailtext.Text = reader("email").ToString
                    htx_KMcattext.Text = reader("KMcat").ToString

                End If
                If Not reader.IsClosed Then
                    reader.Close()
                End If
                comm.Dispose()
            Catch ex As Exception
                Response.Write(ex.ToString)
                Response.End()
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "uploadfail", "<script>alert('檢查帳號錯誤');history.back();</script>")
                Exit Sub
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If
    End Sub

    Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

        Dim account As String, realname As String, nickname As String, password1 As String, password2 As String
        Dim idn As String, member_org As String, com_tel As String, com_ext As String, ptitle As String
        Dim birth As String, birthyear As String, birthmonth As String, birthday As String, sex As String
        Dim homeaddr As String, zip As String, phone As String, home_ext As String, mobile As String, fax As String
        Dim email As String, htx_KMcat As String, htx_KMcatID As String, htx_KMautoID As String

        account = accounttext.Text.Trim
        realname = realnametext.Text.Trim
        nickname = nicknametext.Text.Trim
        password1 = password1text.Text.Trim
        password2 = password2text.Text.Trim
        idn = idntext.Text.Trim
        member_org = member_orgtext.Text.Trim
        com_tel = com_teltext.Text.Trim
        com_ext = com_exttext.Text.Trim
        ptitle = ptitletext.Text.Trim
        birthyear = birthYeartext.Text.Trim
        birthmonth = birthMonthtext.SelectedValue
        birthday = birthdaytext.SelectedValue
        birth = birthyear & birthmonth & birthday
        sex = ""
        If maletext.Checked Then
            sex = "1"
        ElseIf femaletext.Checked Then
            sex = "0"
        End If
        homeaddr = homeaddrtext.Text.Trim
        zip = ziptext.Text.Trim
        phone = phonetext.Text.Trim
        home_ext = home_exttext.Text.Trim
        mobile = mobiletext.Text.Trim
        fax = faxtext.Text.Trim
        email = emailtext.Text.Trim
        htx_KMcat = htx_KMcattext.Text.Trim
        htx_KMcatID = htx_KMcatIDtext.Value
        htx_KMautoID = htx_KMautoIDtext.Value
     
        Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim conn As SqlConnection
        Dim comm As SqlCommand        
        Dim sql As String

        '---check this account is used or not---
        conn = New SqlConnection(ConnString)
        Try
            conn.Open()
            If password1 <> "" Then
                sql = "UPDATE Member SET passwd = @password, nickname = @nickname, birthday = @birthday, sex = @sex, homeaddr = @homeaddr, zip = @zip, "
                sql &= "phone = @phone, home_ext = @home_ext, mobile = @mobile, fax = @fax, email = @email, member_org = @member_org, "
                sql &= "com_tel = @com_tel, com_ext = @com_ext, ptitle = @ptitle WHERE account = @account"
            Else
                sql = "UPDATE Member SET  nickname = @nickname, birthday = @birthday, sex = @sex, homeaddr = @homeaddr, zip = @zip, "
                sql &= "phone = @phone, home_ext = @home_ext, mobile = @mobile, fax = @fax, email = @email, member_org = @member_org, "
                sql &= "com_tel = @com_tel, com_ext = @com_ext, ptitle = @ptitle WHERE account = @account"
            End If
            

            comm = New SqlCommand(sql, conn)
            If password1 <> "" Then
                comm.Parameters.AddWithValue("@password", password1)
            End If
            comm.Parameters.AddWithValue("@nickname", nickname)
            comm.Parameters.AddWithValue("@birthday", birth)
            comm.Parameters.AddWithValue("@sex", sex)
            comm.Parameters.AddWithValue("@homeaddr", homeaddr)
            comm.Parameters.AddWithValue("@zip", zip)
            comm.Parameters.AddWithValue("@phone", phone)
            comm.Parameters.AddWithValue("@home_ext", home_ext)
            comm.Parameters.AddWithValue("@mobile", mobile)
            comm.Parameters.AddWithValue("@fax", fax)
            comm.Parameters.AddWithValue("@email", email)
            comm.Parameters.AddWithValue("@member_org", member_org)
            comm.Parameters.AddWithValue("@com_tel", com_tel)
            comm.Parameters.AddWithValue("@com_ext", com_ext)
            comm.Parameters.AddWithValue("@ptitle", ptitle)
            comm.Parameters.AddWithValue("@account", account)
            comm.ExecuteNonQuery()
            comm.Dispose()
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "accountduplicate", "<script>alert('修改成功');</script>")
            Exit Sub

        Catch ex As Exception
            Response.Write(ex.ToString)
            Response.End()
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "uploadfail", "<script>alert('更新失敗');history.back();</script>")
            Exit Sub
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub


End Class
