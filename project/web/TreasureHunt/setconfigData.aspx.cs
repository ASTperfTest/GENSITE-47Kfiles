using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

public partial class TreasureHunt_setconfigData : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    private int activityId = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (WebUtility.GetStringParameter("tresduremima", string.Empty) != "nungglneiqdesffhfuujvfgk")
            Response.End();
        treasureHunt = new TreasureHunt("");
        if (!IsPostBack)
        {
            activityList.DataSource = CreateDataSource();
            activityList.DataTextField = "TextField";
            activityList.DataValueField = "ValueField";
            activityList.DataBind();
            activityList.SelectedIndex = 0;
        }
        activityId = Convert.ToInt32(activityList.SelectedValue);
        treasureHunt.SetActivity(activityId);

        if (Request.Form.GetValues("saveconfigData") != null && Request.Form.GetValues("saveconfigData")[0].ToString().CompareTo("true") == 0)
            SaveConfigData();

        SetTextData();
    }

    public void SetTextData()
    {
        TextBoxChatModeStartTime.Text = treasureHunt.getCheatMode;
        TextBoxChatModeEndTime.Text = treasureHunt.getCheatModeEnd;
        string tempLotteryStartDate = "";
        string tempLotteryStartHoue="";
        string tempLotteryEndDate = "";
        string tempLotteryEndHoue = "";
        string temp = treasureHunt.getLotteryStartDate;
        if(!string.IsNullOrEmpty(temp))
        {
            tempLotteryStartDate = Convert.ToDateTime(temp).ToString("yyyy/MM/dd");
            tempLotteryStartHoue = Convert.ToDateTime(temp).ToString("HH:mm");
        }
        temp = treasureHunt.getLotteryEndDate;
        if (!string.IsNullOrEmpty(temp))
        {
            tempLotteryEndDate = Convert.ToDateTime(temp).ToString("yyyy/MM/dd");
            tempLotteryEndHoue = Convert.ToDateTime(temp).ToString("HH:mm");
        }
        int selectIndex = 0;
        TextVoteStartDate.Text = tempLotteryStartDate;
        TextBoxTxtVoteStartHours.DataSource = CreateHoursSource(tempLotteryStartHoue,ref selectIndex);
        TextBoxTxtVoteStartHours.DataTextField = "TextField";
        TextBoxTxtVoteStartHours.DataValueField = "ValueField";
        TextBoxTxtVoteStartHours.DataBind();
        TextBoxTxtVoteStartHours.SelectedIndex = selectIndex;
        selectIndex = 0;
        TextBoxTxtVoteEndHours.DataSource = CreateHoursSource(tempLotteryEndHoue, ref selectIndex);
        TextBoxTxtVoteEndHours.DataTextField = "TextField";
        TextBoxTxtVoteEndHours.DataValueField = "ValueField";
        TextBoxTxtVoteEndHours.DataBind();
        TextBoxTxtVoteEndHours.SelectedIndex = selectIndex;
        TextVoteEndDate.Text = tempLotteryEndDate;
    }

    private void SaveConfigData()
    {
        string cheatMode = TextBoxChatModeStartTime.Text;
        string sheatModeEnd = TextBoxChatModeEndTime.Text;
        string startVoteDate = TextVoteStartDate.Text;
        string endVoteDate = TextVoteEndDate.Text ;
        string startVotehours=TextBoxTxtVoteStartHours.SelectedValue;
        string endVotehours = TextBoxTxtVoteEndHours.SelectedValue;
        DateTime dt = new DateTime();
        string message = "";
        bool flag = false;
        if (!string.IsNullOrEmpty(cheatMode) && !string.IsNullOrEmpty(sheatModeEnd))
        {
            if (DateTime.TryParseExact(cheatMode, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out dt) && DateTime.TryParseExact(sheatModeEnd, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out dt))
            {
                treasureHunt.getCheatMode = cheatMode;
                treasureHunt.getCheatModeEnd = sheatModeEnd;
                message = "<script>alert(\"修改成功!!\");</script>";
                flag = true;
            }
        }
        else
        {
            message = "<script>alert(\"備援時間修改失敗!!請檢查輸入格式(需要同時輸入起始結束日)\");</script>";
        }
        if (!string.IsNullOrEmpty(startVoteDate + startVotehours) && !string.IsNullOrEmpty(endVoteDate + " " + endVotehours))
        {
            if (DateTime.TryParseExact(startVoteDate + " " + startVotehours, "yyyy/MM/dd HH:mm", null, System.Globalization.DateTimeStyles.None, out dt)
                && DateTime.TryParseExact(endVoteDate + " " + endVotehours, "yyyy/MM/dd HH:mm", null, System.Globalization.DateTimeStyles.None, out dt))
            {
                treasureHunt.getLotteryStartDate = startVoteDate + " " + startVotehours;
                treasureHunt.getLotteryEndDate = endVoteDate + " " + endVotehours;
                if (!flag)
                    message = "<script>alert(\"修改成功!!\");</script>";
            }
        }
        else
        {
            message = "<script>alert(\"投套數區間修改失敗!!請檢查輸入格式(需要同時輸入起始結束日)\");</script>";
        }
       
        Response.Write(message);
    }

    ICollection CreateDataSource()
    {

        // Create a table to store data for the DropDownList control.
        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("TextField", typeof(String)));
        dt.Columns.Add(new DataColumn("ValueField", typeof(int)));
        IList activityLists = treasureHunt.GetAllActivityList();
        foreach (TreasureHunt.Treasure temp in activityLists)
        {
            dt.Rows.Add(temp.TreasureName,temp.TreasureId);
        }
        // Populate the table with sample values.

        // Create a DataView from the DataTable to act as the data source
        // for the DropDownList control.
        DataView dv = new DataView(dt);
        return dv;

    }

    ICollection CreateHoursSource(string hourss ,ref int selectIndex)
    {

        // Create a table to store data for the DropDownList control.
        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("TextField", typeof(String)));
        dt.Columns.Add(new DataColumn("ValueField", typeof(String)));
        DateTime dt1 = Convert.ToDateTime("2000/01/01");
        DateTime dt2 = Convert.ToDateTime("2000/01/02");
        int i = 0;
        while ((DateTime.Compare(dt2, dt1) > 0))
        {
            dt.Rows.Add(CreateRow(dt1.ToString("HH:mm"), dt1.ToString("HH:mm"),dt));
            if (dt1.ToString("HH:mm").CompareTo(hourss)==0)
                selectIndex = i;
            dt1=dt1.AddHours(1);
            i++;
        }
        
        // Populate the table with sample values.

        // Create a DataView from the DataTable to act as the data source
        // for the DropDownList control.
        DataView dv = new DataView(dt);
        return dv;

    }

    DataRow CreateRow(String Text, int Value, DataTable dt)
    {

        // Create a DataRow using the DataTable defined in the 
        // CreateDataSource method.
        DataRow dr = dt.NewRow();
        // This DataRow contains the ColorTextField and ColorValueField 
        // fields, as defined in the CreateDataSource method. Set the 
        // fields with the appropriate value. Remember that column 0 
        // is defined as ColorTextField, and column 1 is defined as 
        // ColorValueField.
        dr[0] = Text;
        dr[1] = Value;

        return dr;

    }
    DataRow CreateRow(String Text, String Value, DataTable dt)
    {

        // Create a DataRow using the DataTable defined in the 
        // CreateDataSource method.
        DataRow dr = dt.NewRow();
        // This DataRow contains the ColorTextField and ColorValueField 
        // fields, as defined in the CreateDataSource method. Set the 
        // fields with the appropriate value. Remember that column 0 
        // is defined as ColorTextField, and column 1 is defined as 
        // ColorValueField.
        dr[0] = Text;
        dr[1] = Value;

        return dr;

    }
}
