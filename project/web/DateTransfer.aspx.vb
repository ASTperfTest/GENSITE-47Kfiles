Public Partial Class DateTransfer
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            hidDays.Value = DateTime.DaysInMonth(Now.Year, Now.Month)
            txtYear.Text = DateTime.Now.Year
            ddlMonth.SelectedValue = DateTime.Now.Month
            setddlDayValue()
            ddlDay.SelectedValue = DateTime.Now.Day
        End If
    End Sub

    Protected Sub ddlMonth_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMonth.SelectedIndexChanged
        setddlDayValue()
    End Sub

    Private Sub setddlDayValue()
        ddlDay.Items.Clear()
        For i As Integer = 1 To Convert.ToInt16(hidDays.Value)
            ddlDay.Items.Add(i)
        Next i
    End Sub

    Protected Sub rdoCountry_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoCountry.CheckedChanged
        If rdoCountry.Checked = True Then
            setddlDayValue()
        End If
    End Sub

    Protected Sub rdoLunar_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoLunar.CheckedChanged
        If rdoLunar.Checked = True Then
            setddlDayValue()
        End If
    End Sub
End Class