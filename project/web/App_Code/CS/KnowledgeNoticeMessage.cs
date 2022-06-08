using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Web.Configuration;

public class KnowledgeNoticeMessage : MailMessage
{
    public static void SendMessage(int questionId, string poster, string postedMsg, string reDirectURL)
    {
        return;
        try
        {
            SmtpClient smtp = new SmtpClient();
            KnowledgeNoticeMessage msg = null;
            string qTitle = GetQuestionTitle(questionId);


        
            var receivers = GetReceivers(questionId, poster);
            foreach (var item in receivers)
            {
                try
                {
                    msg = new KnowledgeNoticeMessage(questionId, qTitle, postedMsg, reDirectURL);
                    msg.To.Add(item);
                    smtp.Send(msg);
                }
                catch { }
            }

            msg = new KnowledgeNoticeMessage(questionId, qTitle, postedMsg, reDirectURL);
            msg.Bcc.Add(new MailAddress("bob_lin@gss.com.tw"));
            //msg.Bcc.Add(new MailAddress("km@mail.coa.gov.tw"));
            smtp.Send(msg);
            
        }
        catch { }
    }

    private static List<MailAddress> GetReceivers(int questionId, string poster)
    {
        List<MailAddress> r = new List<MailAddress>(3);

        string ConnString = WebConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;
        string sql = @"
        
        
        select
	         isnull(Member.nickname, '') as nickname
	        ,isnull(Member.account, '') as account
	        ,isnull(Member.email, '') as email
        from
        (
	        SELECT   
		        distinct
		        iEditor
	        FROM dbo.CuDTGeneric 
	        INNER JOIN dbo.KnowledgeForum ON dbo.CuDTGeneric.iCUItem = dbo.KnowledgeForum.gicuitem
	        WHERE     
		        (CuDTGeneric.iCUItem = @questionId)
	        or   CuDTGeneric.iCUItem in (
		        select gicuitem from KnowledgeForum where ParentIcuitem = @questionId
		        union
		        select gicuitem from KnowledgeForum where ParentIcuitem in(select gicuitem from KnowledgeForum where ParentIcuitem = @questionId)
	        )
        ) as editors 
        inner join Member on Member.account = editors.iEditor
        where iEditor != @poster ";


        using (SqlConnection conn = new SqlConnection(ConnString))
        {
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("questionId", questionId);
            cmd.Parameters.AddWithValue("poster", poster);

            var myReader = cmd.ExecuteReader();
            while (myReader.Read())
            {

                try
                {
                    var addr = new MailAddress(myReader["email"].ToString(), myReader["nickname"].ToString());
                    r.Add(addr);
                }
                catch { }
            }
        }

        return r;
    }

    private static string GetQuestionTitle(int questionId)
    {
        string ConnString = WebConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;
        string sql = @"select sTitle from CuDTGeneric where iCUItem = @questionId";


        using (SqlConnection conn = new SqlConnection(ConnString))
        {
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("questionId", questionId);

            var title = cmd.ExecuteScalar();

            return title == null ? "" : title.ToString();
        }
    }


    public KnowledgeNoticeMessage(int questionId, string qTitle, string postedMsg, string reDirectURL)
    {
        this.Subject = "農業知識入口網-知識家新訊息通知信";
        this.IsBodyHtml = true;
        this.BodyEncoding = System.Text.Encoding.UTF8;
        this.SubjectEncoding = System.Text.Encoding.UTF8;
        this.Body = GetContent(questionId, qTitle, postedMsg, reDirectURL);
    }

    private string GetContent(int questionId, string qTitle, string postedMsg, string reDirectURL)
    {
        string body = @"
<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"" >
<head>
    <title>農業知識入口網-知識家回覆通知信</title>
    <style type=""text/css"">
        body
        {
            font-size:14px;            
        }          
         
        table {   
          border: 1px solid #4F81BD;   
          border-collapse: collapse;   
        }   
        tr, td {   
          border: 1px solid #00;   
          border-color:#4F81BD;
        }      
        .style1
        {
            font-size:16px;
            height: 20px;
            background-color:#003366; 
            color: #FFFFFF;
        }   
    </style>
</head>
<body>
    <table border=""1"" style=""width:90%"" align='center'>
        <tr>
            <td class=""style1"">農業知識入口網</td>
        </tr>
        <tr>
            <td>您好，<br />
            知識家問題 [<a href=""%URL%"" target='_blank' style=""COLOR: #3b5998; TEXT-DECORATION: none"" target=""_blank"">%title%</a>]，有新的專家回應訊息：<br /><br />
            <table style=""background-color:rgb(252,249,215); width:95%"" align='center'>
                <tr>
                    <td>%msg%<br/></td>
                </tr>
            </table>
            <br />    
            ※ 此信為系統自動發送，請勿回覆<br />
            農業知識入口網 敬上 
            </td>
        </tr>
    </table>
</body>
</html>";

        if (string.IsNullOrEmpty(reDirectURL))
        {
            body = body.Replace("%URL%", WebConfigurationManager.AppSettings["myURL"] + "/knowledge/knowledge_cp.aspx?ArticleId=" + questionId.ToString() + "&ArticleType=A&CategoryId=A&kpi=0");
        }
        else
        {
            body = body.Replace("%URL%", WebConfigurationManager.AppSettings["myURL"] + reDirectURL);
        }

        body = body.Replace("%title%", qTitle);
        body = body.Replace("%msg%", postedMsg.Trim().Replace(System.Environment.NewLine, "<br/>"));

        return body;
    }

}

