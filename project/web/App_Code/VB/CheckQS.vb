Imports Microsoft.VisualBasic

Public Class CheckQS
    Shared Sub Check(ByVal QS As String, ByVal myPage As Page)
        '過濾字元包含 "@<>~'!$^*\(){}[]-+; 和 script:
        Dim r As New Regex("[""@<>~'!$^*\\\(\)\{\}\[\]\-+:]|(%22)|(%40)|(%3c)|(%3e)|(%7e)|(%24)|(%5e)|(%5c)|(%7b)|(%7d)|(%5b)|(%5d)|(script%3a)|(script:)")
        Dim QSArray As String()
        QSArray = QS.Split("&")
        Dim i As Integer
        If QS <> "" Then
            Try
                For i = LBound(QSArray) To UBound(QSArray)
                    Dim m As Match = r.Match(QSArray(i).Split("=")(1).ToLower())
                    If m.Success Then
                        myPage.Response.Redirect("Default.aspx")
                        'myPage.ClientScript.RegisterStartupScript(myPage.GetType(), "", "alert('test')", True)
                    End If
                Next
            Catch
                myPage.Response.Redirect("/")
            End Try
        End If
    End Sub
End Class
