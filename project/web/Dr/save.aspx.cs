using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class save : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string key = Request.QueryString["key"];
        string login_id = Request.Form["login_id"];
        string game_key = Request.Form["game_key"];

        Farm2009 farm = new Farm2009(login_id, game_key);

        if (!farm.isMatchedKey(key))
        {
            Response.Write("KeyError");
        }
        else
        {
            string action = Request.Form["action"];
            string current_level = Request.Form["nowlevel"];
            string current_score = Request.Form["nowScore"];
            string current_timer = Request.Form["nowtime"];
            string current_round = Request.Form["nowround"];

            if (action == "set" && (current_level == null || current_score == null || current_timer == null || current_score=="0" ))
            {
                Response.Write("DataError");
            }
            else
            {
                if (action.Trim() == "set")
                {
                	int save_tmp = 0;
                	
                	 for (int i = 0; i < 10; i++)
                    {
                        string q_id = Request.Form["qa" + (i + 1) + "id"];
                        string q_is_correct = Request.Form["qa" + (i + 1) + "ans"];
                        if (q_id != null && q_is_correct != null)
                        {
                            // 有資料，計數器增加
                            save_tmp++;
                        }
                    }
                    if(save_tmp>=7){
					 int questionId = Convert.ToInt32(Request.Form["qa1id"]);
					 if(questionId<=2143)
					 {
						current_level ="1";
					 }else if(questionId<=2321)
					 {
						current_level ="2";
					 }else if(questionId>2321)
					 {
						current_level ="3";
					 }
                    // 準備存檔紀錄分數
                    int data_id = farm.updateGameData(current_round, current_level, current_score, current_timer);

                    // 2010.2.24 修正
                    // 如果data_id沒有大於零，代表傳入的資料有問題，不寫入question_log
                    if (data_id > 0)
                    {
                        // 紀錄答題紀錄
                        for (int i = 0; i < 10; i++)
                        {
                            string q_id = Request.Form["qa" + (i + 1) + "id"];
                            string q_is_correct = Request.Form["qa" + (i + 1) + "ans"];
                            if (q_id != null && q_is_correct != null)
                            {
                                // 有資料，準備寫入
                                farm.updateQuestionLog(data_id, (i + 1), Convert.ToInt32(q_id), Convert.ToInt32(q_is_correct));
                            }
                        }
                    }

                    // 回傳遊戲資料
                    string xml = farm.getGameData();
                    Response.ContentType = "text/xml";
                    Response.Write(xml);
                }else{
                    Response.Write("DataError");	
                }
                
                }
                else
                {
                    // 僅讀取遊戲資料
                    string xml = farm.getGameData();
                    Response.ContentType = "text/xml";
                    Response.Write(xml);
                }
            }
        }

        farm = null;
    }
}
