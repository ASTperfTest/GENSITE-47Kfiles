Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class ProsData_ProsDataAdd
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

        Dim account As String, realname As String, nickname As String, password1 As String, password2 As String
        Dim idn As String, member_org As String, com_tel As String, com_ext As String, ptitle As String
        Dim birth As String, birthyear As String, birthmonth As String, birthday As String, sex As String
        Dim homeaddr As String, zip As String, phone As String, home_ext As String, mobile As String, fax As String
        Dim email As String, htx_KMcat1 As String, htx_KMcatID1 As String, htx_KMautoID1 As String

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
        htx_KMcat1 = htx_KMcat.Text.Trim
        htx_KMcatID1 = htx_KMcatID.Value
        htx_KMautoID1 = htx_KMautoID.Value

        Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim conn As SqlConnection
        Dim comm As SqlCommand
        Dim reader As SqlDataReader
        Dim sql As String

        '---check this account is used or not---
        conn = New SqlConnection(ConnString)
        Try
            conn.Open()
            sql = "SELECT account FROM Member WHERE account = @account"
            comm = New SqlCommand(sql, conn)
            comm.Parameters.AddWithValue("@account", account)
            reader = comm.ExecuteReader
            If reader.HasRows Then
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "accountduplicate", "<script>alert('此帳號已被註冊');history.back();</script>")
                Exit Sub
            End If
            If Not reader.IsClosed Then
                reader.Close()
            End If
            comm.Dispose()
        Catch ex As Exception
            'Response.Write(ex.ToString)
            'Response.End()
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "uploadfail", "<script>alert('檢查帳號失敗');history.back();</script>")
            Exit Sub
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        
        '---upload image file---
        Dim filename As String = ""

        If imgfile.HasFile Then

            Try
                filename = "prosimgfile" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & Now.Millisecond
                filename &= Mid(imgfile.PostedFile.FileName, InStrRev(imgfile.PostedFile.FileName, "."), imgfile.PostedFile.FileName.Length)
                imgfile.PostedFile.SaveAs(Server.MapPath("../Public/Data/" & filename))

            Catch ex As Exception
                'Response.Write(ex.ToString)
                'Response.End()
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "uploadfail", "<script>alert('上傳圖檔失敗');history.back();</script>")

                Exit Sub

            End Try

        End If

        conn = New SqlConnection(ConnString)
        Try
            conn.Open()

            sql = "INSERT INTO Member(account, passwd, realname, homeaddr, phone, mobile, email, createtime, modifytime, zip, home_ext, birthday, "
            sql &= "sex, member_org, id_type1, com_tel, com_ext, fax, create_user, photo, ID, ptitle, actor, KMcat, KMcatID, KMautoID, logincount, "
            sql &= "orderepaper, nickname) VALUES (@account, @passwd, @realname, @homeaddr, @phone, @mobile, @email, GETDATE(),GETDATE(), @zip, @home_ext, "
            sql &= "@birthday, @sex, @member_org, '1', @com_tel, @com_ext, @fax, 'hyweb', @photo, @id, @ptitle, '5', @KMcat, @KMcatID, @KMautoID, 0, 'Y', @nickname)"

            comm = New SqlCommand(sql, conn)

            comm.Parameters.AddWithValue("@account", account)
            comm.Parameters.AddWithValue("@passwd", password1)
            comm.Parameters.AddWithValue("@realname", realname)
            comm.Parameters.AddWithValue("@homeaddr", homeaddr)
            comm.Parameters.AddWithValue("@phone", phone)
            comm.Parameters.AddWithValue("@mobile", mobile)
            comm.Parameters.AddWithValue("@email", email)
            comm.Parameters.AddWithValue("@zip", zip)
            comm.Parameters.AddWithValue("@home_ext", home_ext)
            comm.Parameters.AddWithValue("@birthday", birth)
            comm.Parameters.AddWithValue("@sex", sex)
            comm.Parameters.AddWithValue("@member_org", member_org)
            comm.Parameters.AddWithValue("@com_tel", com_tel)
            comm.Parameters.AddWithValue("@com_ext", com_ext)
            comm.Parameters.AddWithValue("@fax", fax)
            comm.Parameters.AddWithValue("@photo", filename)
            comm.Parameters.AddWithValue("@id", idn)
            comm.Parameters.AddWithValue("@ptitle", ptitle)
            comm.Parameters.AddWithValue("@KMcat", htx_KMcat1)
            comm.Parameters.AddWithValue("@KMcatID", htx_KMcatID1)
            comm.Parameters.AddWithValue("@KMautoID", htx_KMautoID1)
            comm.Parameters.AddWithValue("@nickname", nickname)

            comm.ExecuteNonQuery()
            comm.Dispose()
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "success", "<script>alert('新增專家成功');</script>")
        Catch ex As Exception
            'Response.Write(ex.ToString)
            'Response.End()
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "success", "<script>alert('資料庫錯誤');</script>")
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

End Class
