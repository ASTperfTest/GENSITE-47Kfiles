Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Collections


Partial Class ViewCounter
    Inherits System.Web.UI.Page

    Public UrlRef As String = ""

    Shared htCounter As Hashtable

    Shared htSubjectCounter As Hashtable

    Shared LastUpdateTime As Date = Date.MinValue

    Shared LastSubjectUpdateTime As Date = Date.MinValue

    Public Shared UpdateInterval As Int32 = 5

    Public Shared ClickCount As Int64 = 0

    Public MP As Int32 = 0

    Shared objLocker As New Object

    Public IsShowAll As Boolean = False

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Me.Request("ShowEX") IsNot Nothing Then
            Dim ex As Exception = Me.Session("Last Exception")
            If ex IsNot Nothing Then
                response.write(ex.ToString)
            Else
                response.write("No Exception !!")
            End If
            response.end()
        End If

        Try
            Response.Buffer = True
            Response.ExpiresAbsolute = Now.AddSeconds(-1)
            Response.Expires = 0
            Response.CacheControl = "no-cache"

            If Me.htCounter Is Nothing Then
                SyncLock Me.objLocker
                    If Me.htCounter Is Nothing Then
                        Me.htCounter = New Hashtable
                    End If
                End SyncLock
            End If

            If Me.htSubjectCounter Is Nothing Then
                SyncLock Me.objLocker
                    If Me.htSubjectCounter Is Nothing Then
                        Me.htSubjectCounter = New Hashtable
                    End If
                End SyncLock
            End If

            If Me.LastUpdateTime = Date.MinValue Then
                SyncLock Me.objLocker
                    If Me.LastUpdateTime = Date.MinValue Then
                        LastUpdateTime = now
                    End If
                End SyncLock
            End If

            If Me.LastSubjectUpdateTime = Date.MinValue Then
                SyncLock Me.objLocker
                    If Me.LastSubjectUpdateTime = Date.MinValue Then
                        LastSubjectUpdateTime = now
                    End If
                End SyncLock
            End If

            If Me.Request("ShowAll") IsNot Nothing Then
                IsShowAll = True
                ShowAll()
            End If


            ClickCount += 1

            AddCounter()

            AddSubjectCounter()
        Catch ex As Exception
            Me.Session("Last Exception") = ex

            Return
        End Try

        Me.Session("Last Exception") = Nothing
    End Sub

    Sub AddSubjectCounter()
        If Page.Request.UrlReferrer IsNot Nothing Then
            UrlRef = Page.Request.UrlReferrer.PathAndQuery().tolower

            '主題館計數
            GetMP()

            If htSubjectCounter(MP) Is Nothing Then
                htSubjectCounter(MP) = 1
            Else
                htSubjectCounter(MP) += 1
            End If
        End If

        If Me.LastSubjectUpdateTime.AddSeconds(UpdateInterval) < Now OrElse Me.Request("Update") IsNot Nothing Then
            Dim IsUpdate As Boolean = False
            Dim OldHT As Hashtable = Nothing

            SyncLock Me.objLocker
                If Me.LastSubjectUpdateTime.AddSeconds(UpdateInterval) < Now OrElse Me.Request("Update") IsNot Nothing Then
                    LastSubjectUpdateTime = Now

                    OldHT = Me.htSubjectCounter
                    Me.htSubjectCounter = New Hashtable

                    IsUpdate = True
                End If
            End SyncLock

            If IsUpdate Then
                UpdateSubjectDB(OldHT)
            End If
        End If
    End Sub

    Sub UpdateSubjectDB(ByVal HT As Hashtable)
        'Response.Write("UpdateDB..." & HT.Keys.Count & ", " & ClickCount & "<br>")

        If HT.Keys.Count > 0 Then
            Dim YMD As String = Now.Tostring("yyyy-MM-dd")

            '先 Update


            For Each Key As Int32 In HT.Keys
                '先 Update
                dim UpdateSQL as String = "Update CounterForSubjectByDate Set [ViewCount] = [ViewCount] + " & HT(Key) & " Where YMD = '" & YMD & "' And CtRootId = " & Key

                Dim RC As Int32 = Me.executeSQL(UpdateSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)

                If RC = 0 Then
                    '測試 Insert
                    Try
                        Dim InsertSQL As String = "Insert Into CounterForSubjectByDate(YMD, CtRootId, ViewCount) Values('" & YMD & "', " & Key & ", " & HT(Key) & ")"
                        'Response.Write(InsertSQL)

                        Me.executeSQL(InsertSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)
                    Catch ex As Exception
                        '再試 Update
                        Me.executeSQL(UpdateSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)
                    End Try
                End If
            Next
        End If
    End Sub

    Sub GetMP()
        Dim SP As Int32
        Dim EP As Int32

        UrlRef = Page.Request.UrlReferrer.PathAndQuery().tolower

        SP = UrlRef.IndexOf("mp=")

        If SP > 0 Then
            EP = UrlRef.IndexOf("&", SP)
            If EP = -1 Then
                EP = UrlRef.Length '下一個字元
            End If

            Me.MP = UrlRef.Substring(SP + 3, EP - SP - 3)
        End If

    End Sub

    Sub AddCounter()
        If Page.Request.UrlReferrer IsNot Nothing Then
            UrlRef = Page.Request.UrlReferrer.PathAndQuery().ToLower

            If Me.UrlRef.StartsWith("/mp.asp?") Then
                Me.AddClick("首頁")
                Me.AddUnitClick()
            ElseIf Me.UrlRef.StartsWith("/lp.asp") OrElse Me.UrlRef.StartsWith("/ct.asp") Then
                If Me.UrlRef.IndexOf("ctnode=1580") > -1 Then
                    Me.AddClick("最新消息")
                ElseIf Me.UrlRef.IndexOf("ctnode=4162") > -1 Then
                    Me.AddClick("農業知識拼圖")
                ElseIf Me.UrlRef.IndexOf("ctnode=1581") > -1 Then
                    Me.AddClick("農業與生活")
                ElseIf Me.UrlRef.IndexOf("ctnode=1582") > -1 Then
                    Me.AddClick("新優質農家")
                ElseIf Me.UrlRef.IndexOf("ctnode=1583") > -1 Then
                    Me.AddClick("產銷專欄")
                ElseIf Me.UrlRef.IndexOf("ctnode=1591") > -1 Then
                    Me.AddClick("資源推薦")
                ElseIf Me.UrlRef.IndexOf("ctnode=1584") > -1 Then
                    Me.AddClick("影音專區")
                ElseIf Me.UrlRef.IndexOf("ctnode=1585") > -1 Then
                    Me.AddClick("相關網站")
                End If
                Me.AddUnitClick()
                If Me.UrlRef.StartsWith("/ct.asp") Then
                    Me.AddCuDTGenericClick("")
                End If

            ElseIf Me.UrlRef.IndexOf("/pedia/") > -1 Then
                Me.AddClick("農業小百科")
                Me.AddCuDTGenericClick("aid")
            ElseIf Me.UrlRef.IndexOf("/knowledge/") > -1 Then
                Me.AddClick("農業知識家")
                Me.AddKnowledgeClick()
            ElseIf Me.UrlRef.IndexOf("/category/") > -1 Then
                Me.AddClick("農業知識庫")
            ElseIf Me.UrlRef.StartsWith("/subjectlist.aspx") OrElse Me.UrlRef.IndexOf("/subject/") > -1 Then
                Me.AddClick("主題館")
                Me.AddUnitClick()
                Me.AddCuDTGenericClick("")
            ElseIf Me.UrlRef.IndexOf("/gardening/") > -1 Then 'by derek 2009/12/17
                Me.AddClick("園藝教室")
            ElseIf Me.UrlRef.IndexOf("/fish/") > -1 Then
                Me.AddClick("豆仔水族箱")
            ElseIf Me.UrlRef.IndexOf("/dr/") > -1 Then
                Me.AddClick("農業博識王")
            ElseIf Me.UrlRef.IndexOf("/jigsaw2010/") > -1 Then
                Me.AddClick("農作物地圖")
                If Me.UrlRef.IndexOf("/jigsaw2010/detail.aspx") > -1 Then
                    Me.AddCuDTGenericClick("item")
                End If
            ElseIf Me.UrlRef.IndexOf("/century/") > -1 Then
                Me.AddClick("百年農業發展史")
            End If
        End If

        If Me.LastUpdateTime.AddSeconds(UpdateInterval) < Now OrElse Me.Request("Update") IsNot Nothing Then
            Dim IsUpdate As Boolean = False
            Dim OldHT As Hashtable
            SyncLock Me.objLocker
                If Me.LastUpdateTime.AddSeconds(UpdateInterval) < Now OrElse Me.Request("Update") IsNot Nothing Then
                    LastUpdateTime = Now

                    OldHT = Me.htCounter
                    Me.htCounter = New Hashtable
                    IsUpdate = True
                End If
            End SyncLock

            If IsUpdate Then
                UpdateDB(OldHT)
            End If
        End If
    End Sub

    '    2.2.	最新消息, 農業知識拼圖, 農業與生活, 新優值農家, 產銷專欄, 資源推薦, 影音專區, 相關網站的部份以下程式需修正：
    'http://kmweb.coa.gov.tw/lp.asp
    'http://kmweb.coa.gov.tw/ct.asp
    '以 ctNode 這個參數來判斷計數器, 其中最新消息(ctNode = 1580), 知識拼圖(ctNode = 4162), 農業與生活(ctNode = 1581), 新優質農家(ctNode = 1582), 
    '產銷專欄(ctNode = 1583), 資源推薦(ctNode = 1591), 影音專區(ctNode = 1584), 相關網站(ctNode = 1585)
    '2.3.	農業小百科、農業知識家、農業知識庫建議找到所有頁面的共用元件(ascx), 可在 ascx 中加入計算的 function, 並以網址判斷，
    '網址中有包括/Pedia/的則為小百科，有 /knowledge/ 則為知識家, 有/CatTree/的則為知識庫。若沒有共用的 ascx, 最少要在以下的頁面加入 Counter 的 function. 

    Sub UpdateDB(ByVal HT As Hashtable)
        'Response.Write("UpdateDB..." & HT.Keys.Count & ", " & ClickCount & "<br>")

        If HT.Keys.Count > 0 Then
            Dim YMD As String = Now.Tostring("yyyy-MM-dd")

            '先 Update
            Dim UpdateSQL As String = "Update CounterByDate Set "

            For Each Key As String In HT.Keys
                UpdateSQL &= "[" & Key & "] = [" & Key & "] + " & HT(Key) & ", "
            Next

            UpdateSQL = UpdateSQL.Substring(0, UpdateSQL.Length - 2)

            UpdateSQL &= " Where YMD = '" & YMD & "'"


            Dim RC As Int32 = Me.executeSQL(UpdateSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)

            If Me.IsShowAll Then Response.Write(UpdateSQL)

            If Me.IsShowAll Then Response.Write("RC: " & RC & "<br>")

            If Me.IsShowAll Then response.write(WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)

            If RC = 0 Then
                '測試 Insert
                Try
                    Dim InsertSQL As String = "Insert Into CounterByDate("

                    For Each Key As String In HT.Keys
                        InsertSQL &= "[" & Key & "], "
                    Next

                    InsertSQL &= "YMD) Values("

                    For Each Key As String In HT.Keys
                        InsertSQL &= HT(Key) & ", "
                    Next

                    InsertSQL &= "'" & YMD & "')"

                    If Me.IsShowAll Then Response.Write(InsertSQL)

                    RC = Me.executeSQL(InsertSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)

                    If Me.IsShowAll Then Response.Write("RC: " & RC & "<br>")
                Catch ex As Exception
                    '再試 Update
                    RC = Me.executeSQL(UpdateSQL, WebConfigurationManager.ConnectionStrings("GSSConnString").ConnectionString)

                    If Me.IsShowAll Then Response.Write(UpdateSQL)

                    If Me.IsShowAll Then Response.Write("RC: " & RC & "<br>")
                End Try
            End If
        End If
    End Sub

    Sub AddClick(ByVal Key As String)
        If htCounter(Key) Is Nothing Then
            htCounter(Key) = 1
        Else
            htCounter(Key) += 1
        End If
    End Sub
#Region "// sam 2011/05/11 count for Article and mp site "

    Sub AddUnitClick()
        Dim myConnection As SqlConnection = New SqlConnection(WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Try
            Dim mpId As String = GetParameter(Me.UrlRef.ToString(), "mp")
            Dim sql As String = ""
            sql = " if exists(select * from counter where mp =@mp) "
            sql += " begin "
            sql += "        UPDATE counter SET counts = counts + 1  WHERE mp=@mp "
            sql += " End"
            sql += "    else begin "
            sql += "        INSERT INTO counter (mp, counts) VALUES (@mp,1) "
            sql += " End "
            myConnection.Open()
            Dim myCommand = New SqlCommand(sql, myConnection)
            myCommand.Parameters.AddWithValue("@mp", mpId)
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
        Catch ex As Exception
            
        Finally
            myConnection.Close()
        End Try

    End Sub

    Function GetParameter(ByVal queryString As String, ByVal parameterName As String) As String
        parameterName = parameterName & "="
        If queryString.Length > 0 Then
            Dim begin As Integer = queryString.IndexOf(parameterName)
            Dim ends As Integer = 0
            If begin <> -1 Then
                begin += parameterName.Length
                ends = queryString.IndexOf("&", begin)

                If ends = -1 Then
                    ends = queryString.Length
                End If
                Return Server.UrlEncode(queryString.Substring(begin, ends - begin))
            End If
        End If
        Return "null"
    End Function

    Sub AddCuDTGenericClick(ByVal queryParam As String)
        Dim myConnection As SqlConnection = New SqlConnection(WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Try
            Dim cuItem As String = ""
            If queryParam <> "" Then
                cuItem = GetParameter(Me.UrlRef.ToString(), queryParam)
            Else
                Dim ctNodeId = GetParameter(Me.UrlRef.ToString(), "ctnode")
                If ctNodeId = "null" Then
                    cuItem = GetParameter(Me.UrlRef.ToString(), "cuitem")
                Else
                    cuItem = GetParameter(Me.UrlRef.ToString(), "xitem")
                End If
            End If
            If cuItem = "null" Then
                Exit Sub
            End If
            Dim sql As String = "UPDATE CuDTGeneric SET ClickCount = ClickCount + 1 WHERE iCUItem =  @iCUItem "
            myConnection.Open()
            Dim myCommand = New SqlCommand(sql, myConnection)
            myCommand.Parameters.AddWithValue("@iCUItem", cuItem)
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
            Me.AddDailyClick(cuItem)
        Catch ex As Exception
            
        Finally
            myConnection.Close()
        End Try
    End Sub
    Sub AddKnowledgeClick()
        Dim myConnection As SqlConnection = New SqlConnection(WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Try
            Dim articleId = GetParameter(Me.UrlRef.ToString(), "articleid")
            If articleId = "null" Then
                Exit Sub
            End If
            Dim sql As String = "UPDATE KnowledgeForum SET BrowseCount = BrowseCount + 1 WHERE gicuitem = @iCUItem"
            sql += " UPDATE CuDTGeneric SET ClickCount = (SELECT BrowseCount FROM KnowledgeForum WHERE gicuitem = @iCUItem) WHERE icuitem = @iCUItem "
            myConnection.Open()
            Dim myCommand = New SqlCommand(sql, myConnection)
            myCommand.Parameters.AddWithValue("@iCUItem", ArticleId)
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
            Me.AddDailyClick(articleId)
        Catch ex As Exception
            
        Finally
            myConnection.Close()
        End Try
    End Sub

    Sub AddDailyClick(ByVal icuitem As String)
        Dim myConnection As SqlConnection = New SqlConnection(WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Try
            If icuitem = "null" Or icuitem = "" Then
                Exit Sub
            End If
            Dim sql As String = ""
            sql += " if EXISTS(SELECT * FROM DailyClick where iCUItem = @iCUItem and DATEDIFF(Day,editDate,GETDATE()) = 0 ) "
            sql += " begin "
            sql += " UPDATE DailyClick SET dailyClick = dailyClick+1 where iCUItem =  @iCUItem and DATEDIFF(Day,editDate,GETDATE()) = 0 "
            sql += " End "
            sql += " else begin "
            sql += " INSERT INTO DailyClick (iCUItem,dailyClick,editDate) VALUES ( @iCUItem ,'1' , GETDATE() ) "
            sql += " End "
            myConnection.Open()
            Dim myCommand = New SqlCommand(sql, myConnection)
            myCommand.Parameters.AddWithValue("@iCUItem", icuitem)
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
        Catch ex As Exception
            
        Finally
            myConnection.Close()
        End Try
    End Sub
#End Region

    Sub ShowAll()
        Response.Write("LastUpdateTime: " & Me.LastUpdateTime & "<br>")
        Response.Write("ShowAll..." & Me.htCounter.Keys.Count & ", " & ClickCount & "<br>")

        For Each key As String In Me.htCounter.Keys
            Response.Write(key & ": " & Me.htCounter(key) & "<br>")
        Next

        Response.Write("LastSubjectUpdateTime: " & Me.LastSubjectUpdateTime & "<br>")
        Response.Write("Subject...<br>")
        For Each key As Int32 In Me.htSubjectCounter.Keys
            Response.Write(key & ": " & Me.htSubjectCounter(key) & "<br>")
        Next
    End Sub

    Public Shared Function executeSQL(ByVal SQL As String, ByVal DBC As String) As Int32

        Dim sqlConn As New SqlConnection(DBC)
        Dim RC As Int32 = 0
        Try
            sqlConn.Open()

            Dim Cmd As SqlCommand = New SqlCommand
            With Cmd
                .CommandText = SQL
                .CommandType = CommandType.Text
                .Connection = sqlConn
                RC = .ExecuteNonQuery()
            End With
        Catch ex As Exception
        Finally
            sqlConn.Close()
        End Try
        Return RC
    End Function
End Class

'[首頁] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_首頁]  DEFAULT ((0)),
'	[最新消息] [bigint] NOT NULL CONSTRAINT [DF_Table_1_最新消]  DEFAULT ((0)),
'	[農業知識拼圖] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_農業知識拼圖]  DEFAULT ((0)),
'	[農業與生活] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_農業與生活]  DEFAULT ((0)),
'	[新優值農家] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_新優值農家]  DEFAULT ((0)),
'	[產銷專欄] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_產銷專欄]  DEFAULT ((0)),
'	[資源推薦] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_資源推薦]  DEFAULT ((0)),
'	[影音專區] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_影音專區]  DEFAULT ((0)),
'	[相關網站] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_相關網站]  DEFAULT ((0)),
'	[農業小百科] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_農業小百科]  DEFAULT ((0)),
'	[農業知識家] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_農業知識家]  DEFAULT ((0)),
'	[農業知識庫] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_農業知識庫]  DEFAULT ((0)),
'	[主題館] [bigint] NOT NULL CONSTRAINT [DF_CounterByDate_主題館]  DEFAULT ((0)),
