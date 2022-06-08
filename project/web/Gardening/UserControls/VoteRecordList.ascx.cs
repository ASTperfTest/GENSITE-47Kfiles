using System;
using System.Collections;
using System.Data;
using System.Web.UI;

public partial class UserControls_VoteRecordList : UserControl
{
    private DataTable source;
    private string ownerId;

    public string OwnerId
    {
        get
        {
            return ownerId;
        }
        set
        {
            ownerId = value;
        }
    }

    public DataTable Source
    {
        set
        {
            source = value;
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        GridView1.DataSource = source;
        GridView1.DataBind();
    }

    protected void HideRecord_Click(object sender, EventArgs e)
    {
        this.Visible = false;
    }
}
