using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class jigsaw10_PhoneticNotation : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ShowData();           
    }

    protected void Repeater_OnItemCommand(object source, RepeaterCommandEventArgs e)
    {
        try
        {
            string sword = e.CommandArgument.ToString();
            mGIPcoanewDataContext db = new mGIPcoanewDataContext();
            if (e.CommandName == "Delete")
            {                
                PhoneticNotation PhoneticNotationRecord =
                    db.PhoneticNotation.First(c => c.word.ToString().Equals(sword));

                db.PhoneticNotation.DeleteOnSubmit(PhoneticNotationRecord);
                db.SubmitChanges();
                ShowData();
            }
            if (e.CommandName == "Update")
            {
                PhoneticNotation PhoneticNotationRecord =
                    db.PhoneticNotation.First(c => c.word.ToString().Equals(sword));
                txtWord.Text = PhoneticNotationRecord.word.ToString();
                txtWord.ReadOnly = true;
                bsave.Visible = false;
                bUpdate.Visible = true;
                bCancel.Visible = true;
                txtphonetec.Text = PhoneticNotationRecord.PhoneticNotation1.ToString();
                
            }
            
        }
        catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); Response.End(); }


    }
    protected void ShowData()
    {
        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        var result = from p in _mGIPcoanew_repository.List<PhoneticNotation>()
                     select p;
        rptList.DataSource = result;
        rptList.DataBind();
        bsave.Visible = true;
        bUpdate.Visible = false;
        bCancel.Visible = false;
        txtWord.ReadOnly = false;
        
    }

    protected void bsave_Click(object sender, EventArgs e)
    {
        if (checkdata())
        {
            try
            {
                mGIPcoanewDataContext db = new mGIPcoanewDataContext();
                int icount = (from d in db.PhoneticNotation.AsEnumerable()
                              where d.word.ToString().Contains(txtWord.Text)
                              select d).Count();
                if (icount > 0)
                {
                    Page.ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('資料已存在');", true);
                }
                else
                {
                    IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());
                    char[] ss = txtWord.Text.ToCharArray();
                    PhoneticNotation PhoneticNotationRecord = new PhoneticNotation
                    {
                        word = ss[0],
                        PhoneticNotation1 = txtphonetec.Text
                    };
                    _mGIPcoanew_repository.Create<PhoneticNotation>(PhoneticNotationRecord);
                    ShowData();
                    txtWord.Text = "";
                    txtphonetec.Text = "";
                }                
            }
            catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); Response.End(); }            
        }        
    }
    protected bool checkdata()
    {

        if (string.IsNullOrEmpty(txtWord.Text) || string.IsNullOrEmpty(txtphonetec.Text))
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('欄位不可為空白');", true);
            return false;
        }
        char[] ss = txtWord.Text.ToCharArray();
        if (ss.Count() > 1)
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('只能輸入單一字');", true);
            return false;
        }                   
        
        return true;
    }
    protected void bUpdate_Click(object sender, EventArgs e)
    {
        if (checkdata())
        {
            mGIPcoanewDataContext db = new mGIPcoanewDataContext();

            var result = (from p in db.PhoneticNotation.AsEnumerable()
                         where p.word.ToString().Contains(txtWord.Text)
                         select p).First();
            if (result != null)
            {                
                result.PhoneticNotation1 = txtphonetec.Text;
                db.SubmitChanges();
            }
            ShowData();
            txtWord.Text = "";
            txtphonetec.Text = "";
        }        
    }
    protected void bCancel_Click(object sender, EventArgs e)
    {
        txtWord.Text = "";
        txtphonetec.Text = "";
        bsave.Visible = true;
        bUpdate.Visible = false;
        bCancel.Visible = false;
        txtWord.ReadOnly = false;
    }
}
