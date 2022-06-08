using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Net.Mail;
using System.Collections;

public partial class MemberInvitePage : System.Web.UI.Page
{
    protected string strURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"].ToString();
    string Script = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/mp.asp';</script>";
    string sTemplate = "{0}/Member/MIP.aspx?{1}={2}";
    string MemberId;
    string MemberName = string.Empty;
    protected string InvitationCode = string.Empty;
    ArrayList arrTextBox = new ArrayList();
    protected string sendOK = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || string.IsNullOrEmpty(Session["memID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            MemberId = Session["memID"].ToString();
            InitControl();
        }
    }

    protected void InitControl()
    {
        // 會員姓名
        string strQueryScript = @"SELECT realname From Member WHERE account = @account";
        MemberName = SqlHelper.ReturnScalar("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@account", "@account", MemberId)).ToString().Trim();

        litContent.Text = "<textarea name='txtContent' id='txtContent' rows='5' cols='30' readOnly='readonly' style='resize:none;width:300px;margin:10px;'>"
                          + "Hi, 我是" + MemberName + "，介紹你一個好站：農業知識入口網，可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入會員吧~~</textarea>";
       

        // 產生InvitationCode & 顯示於下方URL
        while (string.IsNullOrEmpty(InvitationCode))
        {
            InvitationCode = CheckAccount();
        }

        string shareURL = string.Format(sTemplate, strURL,"Ucode", InvitationCode);

        txtInvitationCode.Text = shareURL;

        // 將panelMail中的TextBox定義在 arrTextBox
        foreach (Control item in panelMail.Controls)
        {
            if (item.GetType() == typeof(TextBox))
            {
                arrTextBox.Add((TextBox)item);
            }
        }

        // 查出以透過該會員加入之會員名單
        string strQueryAccountScript = @"SELECT     M.nickname,M.createtime, I.IsActive 
                                         FROM       member AS M INNER JOIN InviteFriends_Detail AS I
                                                    ON M.account = I.inviteAccount 
                                         WHERE      I.InvitationCode = @InvitationCode
                                                and I.IsActive=1
                                    ";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN",strQueryAccountScript,
            DbProviderFactories.CreateParameter("ODBCDSN","@InvitationCode","@InvitationCode",InvitationCode)))
        {
            if (reader.HasRows)
            {
                int i=0;
                while (reader.Read())
                {
                    i++;
                    string strTemplate = @"<tr>
                                            <td>{0}</td>
                                            <td>{1}</td>
                                            <td>{2}</td>
                                            <td>{3}</td></tr>";

                    string state = string.Empty;
                    switch (reader["IsActive"].ToString())
                    {
                        case "0":
                            state = "未通過";
                            break;
                        case "1":
                            state = "已通過";
                            break;
                        case "9":
                            state = "已被停權";
                            break;
                    }
                    
                    labInvitation.Text += string.Format(strTemplate, i
                        , reader["nickname"].ToString()
                        , reader["createtime"].ToString()
                        , state);
                }
            }
            else
            {
                labInvitation.Text = "<tr><td colspan='10' style='text-align:center;'>尚無邀請好友，快去邀請您的好友加入吧。</td></tr>";
            }
        }
    }

    // 發送邀請信
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        try
        {
            SendInvitationMail(AddMailMessage());
        }
        catch (Exception)
        {
            Response.Write("<script language='javascript' type='text/javascript'>alert('寄送失敗');</script>");
        }

    }

    // Share to Facebook
    protected void btnFBSubmit_Click(object sender, EventArgs e)
    {
        if(!string.IsNullOrEmpty(InvitationCode))
        {
            Response.Redirect(string.Format("{0}/Member/{1}?Scode={2}&v=1", strURL, "MIP.aspx", InvitationCode));
        }
    }

    /// <summary>
    /// 判斷Account是否存在。
    /// Yes: 回傳Code；No: 執行CreateInvitationCode並回傳結果。
    /// </summary>
    /// <returns>string InvitationCode</returns>
    private string CheckAccount()
    {
        string strQueryScript = @"SELECT * FROM InviteFriends_Head WHERE Account = @Account";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@Account", "@Account", MemberId)))
        {
            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    return reader["InvitationCode"].ToString();
                }
                else
                {
                    return CreateInvitationCode();
                }
            }
            else
            {
                return CreateInvitationCode();
            }
        }
    }

    /// <summary>
    /// 創造InvitationCode，並回傳結果。
    /// </summary>
    /// <returns>string InvitationCode</returns>
    private string CreateInvitationCode()
    {
        try
        {
            string strInsertScript = @"DECLARE @rndCode bigint
                                       SET @rndCode = rand() * 100000000000;
                                       INSERT INTO InviteFriends_Head (Account, InvitationCode) 
                                       VALUES (@Account, @rndCode);

                                       select @rndCode";

            string resultCode = SqlHelper.ReturnScalar("ODBCDSN", strInsertScript,
                DbProviderFactories.CreateParameter("ODBCDSN", "@Account", "@Account", MemberId)).ToString();
            return resultCode;
        }
        catch (Exception)
        {
            return string.Empty;
            //throw;
        }
    }

    /// <summary>
    /// 加入Mail的相關資訊
    /// </summary>
    /// <returns>回傳MailMessage</returns>
    protected MailMessage AddMailMessage()
    {
        try
        {

            MailMessage myMessage = new MailMessage();
            MailAddress GSSMail = new MailAddress(System.Web.Configuration.WebConfigurationManager.AppSettings["MailFrom"].ToString());
            string shareURL = string.Format(sTemplate, strURL, "Ecode", InvitationCode);


            myMessage.From = GSSMail;
            foreach (TextBox item in arrTextBox)
            {
                if (!string.IsNullOrEmpty(item.Text))
                {
                    myMessage.Bcc.Add(new MailAddress(item.Text));
                }
            }

            myMessage.Subject = "一起加入農業知識入口網吧";
            myMessage.IsBodyHtml = true;
            //myMessage.Body = string.Format(@"Hi, 我是{0}，介紹你一個好站：<br/>農業知識入口網，可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入會員吧～～<br/>加入會員網址:<br/> <a href='{1}'>{1}</a> <br/><br/>{2}"
            //                    , MemberName, shareURL, MemberName);

            myMessage.Body = GetInvitationString(MemberName, shareURL);

            return myMessage;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 將信件送出。
    /// </summary>
    /// <param name="myMessage">MailMessage</param>
    protected void SendInvitationMail(MailMessage myMessage)
    {
        // SMTP 主機 & 帳號都在Web.Config -> <System.Net><MailSettings>
        try
        {
            if (myMessage != null)
            {
                SmtpClient mySmtp = new SmtpClient();
                mySmtp.Send(myMessage);

                sendOK = "Y";
            }
        }
        catch (Exception)
        {
            throw;
        }
    }



    private string GetInvitationString(string memberName, string shareURL)
    {
        const string body = @"
            <html>
            <head>
            <meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
            <title>農業知識入口網-邀請信</title>
            <style type='text/css'>
            body {
	            background-image: url({baseHref}/images/invitation/bg.jpg);
            }
            table{
	            border:1px solid #690;
            }
            </style>
            </head>
            <body>
            <table width='640' border='0' align='center' cellpadding='0' cellspacing='0'  background='{baseHref}/images/invitation/background.gif'>
              <tr>
                <td width='40' >&nbsp;</td>
                <td width='600' height='480' valign='top'>
                    <br/><br/><br/><br/><br/><br/>
    	            <p>哈囉！我是{memberName}，<br />介紹你一個好站：
			            農業知識入口網，它可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入會員吧~~<br />
			            <br />加入會員網址：<br />
			            <a href='{shareURL}' target='_blank'>{shareURL}</a>
		            </p>
                    <p>{memberName}</p>
                </td>
              </tr>
            </table>
            </body>
            </html>";

        string hostUrl = "http://" + Request.Url.Host;


        return body.Replace("{baseHref}", hostUrl)
            .Replace("{memberName}", memberName)
            .Replace("{shareURL}", shareURL);
    }

}