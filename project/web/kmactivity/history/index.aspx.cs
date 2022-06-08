using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class kmactivity_history_index : System.Web.UI.Page
{
    protected string imageUrl = "";
    protected string memid = "";
    protected string hispicarea = "true";
    protected string openPicture = "false";
    protected string openNode = "";
    protected string canAnswer = "true";
    protected string kmurl = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"].ToString();
    private HistoryPicture historyPicture;
    protected string hideGuid = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || Session["memID"].ToString() == "")
        {
            Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');window.location.href=\"/kmactivity/history/01.aspx?a=history\";</script>");
            Response.End();
        }
        memid = Session["memID"].ToString().Trim();
        historyPicture = new HistoryPicture(memid);

        historyPicture.loginUpdate(memid);
        if (historyPicture.activity != null)
        {
            if (historyPicture.CanNotPlayGame(memid))
            {
                Response.Write("<script>alert('您的遊戲資料有問題，請盡速與網站管理員聯絡!!');window.location.href=\"/kmactivity/history/01.aspx?a=history\";</script>");
                Response.End();
            }
            if (WebUtility.GetStringParameter("hideAnswer", "") != "")
            {
                hideGuid = WebUtility.GetStringParameter("hideguid", "");
                if (string.Compare(hideGuid, Session["hideguid"].ToString(), true) != 0)
                {
                    Response.Write("<script>alert('請勿開兩個視窗玩喔!!');window.location.href=\"/kmactivity/history/01.aspx?a=history\";</script>");
                    Response.End();
                }
                CheckAnswer(WebUtility.GetStringParameter("hideAnswer", ""));
            }
            if (WebUtility.GetStringParameter("hitNode", "") != "")
            {
                HitHint(WebUtility.GetStringParameter("hitNode", ""));
            }
            if (WebUtility.GetStringParameter("NextQuestion", "") != "")
            {
                GetNextQuestion();
            }
            
            imageUrl = "images.aspx?a=" + Guid.NewGuid();
            GenerateQuestion();
            hideGuid = Guid.NewGuid().ToString() ;
            Session["hideguid"] = hideGuid;
        }
    }
    private void GenerateQuestion()
    {
        HistoryPicture.UserQuestionNow userqunow = historyPicture.GetUserQuestionNow();
        if(userqunow != null)
            if (userqunow != null && userqunow.STATES != 1 && userqunow.STATES != 3 && userqunow.STATES != 4)
            {
                DisplayQuestion(userqunow.NodesOrg, userqunow.Nodes, userqunow.STATES, userqunow.CurrentNode);
            }
            else if (userqunow.STATES == 0)
            {
                userqunow = historyPicture.GetUserQuestionNow();
                DisplayQuestion(userqunow.NodesOrg, userqunow.Nodes, userqunow.STATES, userqunow.CurrentNode);
            }
            else if (userqunow.STATES == 4)
            {
                if (historyPicture.GetNewQuestions(historyPicture.userNickname, historyPicture.userRealname, historyPicture.userEmail) == 0)
                {
                    Response.Write("<script>alert('您本日答題超過30題囉!!');</script>");
                    Response.Write("<script>window.location.href=\"/kmactivity/history/top.aspx?a=history\";</script>");
                }
                userqunow = historyPicture.GetUserQuestionNow();
                DisplayQuestion(userqunow.NodesOrg, userqunow.Nodes, userqunow.STATES, userqunow.CurrentNode);
            }
            else
            {
                //if (historyPicture.GetNewQuestions(historyPicture.userNickname, historyPicture.userRealname, historyPicture.userEmail) == 0)
                //{
                //    Response.Write("<script>alert('您本日答題超過30題囉!!');</script>");
                //    Response.Write("<script>window.location.href=\"/\";</script>");
                //}
                userqunow = historyPicture.GetUserQuestionNow();
                DisplayQuestion(userqunow.NodesOrg, userqunow.Nodes, userqunow.STATES, userqunow.CurrentNode);
            }
    }

    private void DisplayQuestion(string nodesOrg,string nodes,int states,int currentNode)
    {
        answerMessageArea.Visible = false;
        tippps.Visible = true;
        string ss = "";
        ss += "<table width=\"100%\" align=\"center\" id=\"historyHintArea\" style=\"border-collapse: separate;border:0px;\" ><tr  style=\"border-collapse: separate;border:0px;text-align:center;\">";
        var dt = historyPicture.GetQuestionById(nodesOrg);
        if (dt.Rows.Count > 0)
        {
            int i = 0;
            if (states != 1 && states != 3)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (i > 1 && ((i % 5) == 0))
                    {
                        ss += "<tr  style=\"border-collapse: separate;border:0px;text-align:center;\">";
                    }
                    //ss += "<td  style=\"border-collapse: separate;border:0px;text-align:left;\">";
                    if (nodes.Contains(dr["ctNodeId"].ToString()))
                    {
                        ss += "<td style=\"border-collapse: separate;border:0px;\"><input  style=\"cursor: pointer;float:left;\" type=\"radio\" name=\"historyanswer\" value=\"" + dr["ctNodeId"].ToString() + "\" /></td><td style=\"border-collapse: separate;border:0px;\"><div style=\"cursor: pointer;float:left;word-wrap:break-word;\" title=\"" + dr["fullPath"].ToString() + "\" onclick=\"hitHint('" + dr["ctNodeId"].ToString() + "')\">" + dr["CatName"].ToString() + "</div></td>";
                    }
                    else
                    {
                        ss += "<td style=\"border-collapse: separate;border:0px;\"><input style=\"float:left;\" type=\"radio\" disabled=\"disabled\" name=\"historyanswer\" value=\"" + dr["ctNodeId"].ToString() + "\" /></td><td style=\"border-collapse: separate;border:0px;\"><div style=\"color:gray;float:left;word-wrap:break-word;\">" + dr["CatName"].ToString() + "</div></td>";
                    }
                    // ss += "</td>";
                    if (i != 0 && (((i + 1) % 5) == 0))
                    {
                        ss += "</tr>";
                    }
                    i++;
                }
            }
            else
            {
                canAnswer = "false";
                if (states == 3)
                {
                    yesnoimage.ImageUrl = "/kmactivity/history/image/yes.gif";
                    answerMessageArea.Visible = true;
                    tippps.Visible = false;
                }
                else if (states == 1)
                {
                    yesnoimage.ImageUrl = "/kmactivity/history/image/no.gif";
                    answerMessageArea.Visible = true;
                    tippps.Visible = false;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    if (i > 1 && ((i % 5) == 0))
                    {
                        ss += "<tr  style=\"border-collapse: separate;border:0px;text-align:center;\">";
                    }
                    //ss += "<td  style=\"border-collapse: separate;border:0px;text-align:left;\">";
                    if (string.Compare(dr["ctNodeId"].ToString(), currentNode.ToString(),true) ==0)
                    {
                        ss += "<td style=\"border-collapse: separate;border:0px;\"><input  style=\"float:left;\" type=\"radio\" disabled=\"disabled\" name=\"historyanswer\" value=\"" + dr["ctNodeId"].ToString() + "\" /></td><td style=\"border-collapse: separate;border:0px;\"><div style=\"color:red;float:left;word-wrap:break-word;\" title=\"" + dr["fullPath"].ToString() + "\">" + dr["CatName"].ToString() + "</div></td>";
                    }
                    else
                    {
                        ss += "<td style=\"border-collapse: separate;border:0px;\"><input style=\"float:left;\" type=\"radio\" disabled=\"disabled\" name=\"historyanswer\" value=\"" + dr["ctNodeId"].ToString() + "\" /></td><td style=\"border-collapse: separate;border:0px;\"><div style=\"color:gray;float:left;word-wrap:break-word;\">" + dr["CatName"].ToString() + "</div></td>";
                    }
                    // ss += "</td>";
                    if (i != 0 && (((i + 1) % 5) == 0))
                    {
                        ss += "</tr>";
                    }
                    i++;
                }
            }
            if (((i % 5) != 0))
                ss += "</tr>";
        }
        ss += "</table>";
        hintstring.Text = ss;
    }

    private void CheckAnswer(string nodeId)
    {
        hispicarea = "false";
        int node = 0;
        if(int.TryParse(nodeId,out node))
        {
            if (historyPicture.CheckAnswer(node))
            {
                yesnoimage.ImageUrl = "/kmactivity/history/image/yes.gif";
                answerMessageArea.Visible = true;
            }
            else
            {
                yesnoimage.ImageUrl = "/kmactivity/history/image/no.gif";
                answerMessageArea.Visible = true;
            }
        }else
        {
            Response.Write("<script>alert('您輸入的資料有問題喔!!');</script>");
            Response.Redirect("/");
            Response.End();
        }
    }

    private void HitHint(string nodeId)
    {
        openPicture = "false";
        int node = 0;
        if (int.TryParse(nodeId, out node))
        {
            if (historyPicture.HitHint(node))
            {
                openPicture = "true";
                openNode = nodeId;
            }
        }
        else
        {
            Response.Write("<script>alert('您輸入的資料有問題喔!!');</script>");
            Response.Redirect("/");
            Response.End();
        }
    }

    private void GetNextQuestion()
    {
        HistoryPicture.UserQuestionNow userqunow = historyPicture.GetUserQuestionNow();
        if (userqunow.STATES == 1 || userqunow.STATES == 3)
        if (historyPicture.GetNewQuestions(historyPicture.userNickname, historyPicture.userRealname, historyPicture.userEmail) == 0)
        {
            Response.Write("<script>alert('您本日答題超過30題囉!!');</script>");
            Response.Write("<script>window.location.href=\"/kmactivity/history/top.aspx?a=history\";</script>");
        }
    }
}