using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using GSS.Vitals.COA.Data;
using Newtonsoft.Json;

public partial class SubjectTest : System.Web.UI.Page
{
    protected string result = string.Empty;

    protected string getmusic = string.Empty;

    protected void Page_Load(object sender, EventArgs e)   
    {
        bool isOK;
        int xItem = 0;
        List<imageItem> totalImagetItem = new List<imageItem>();

        //var buList = 


        List<string> MusicRandom = new List<string>();
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", "set nocount on; select bgMusicID from BackgroundMusic;"))
        {            
            int i = 0;
            while (reader.Read())
            {
                MusicRandom.Add(reader["bgMusicID"].ToString());
            }
        }
        
        
        
        isOK = int.TryParse(Request["xItem"], out xItem);
        if (!isOK) isOK = int.TryParse(Request["CuItem"], out xItem);
        if (isOK)
        {
            string sqlString3 = @"SELECT xiCuItem,aTitle,NFileName,bList,listSeq,bgMusic
                                 FROM CuDtattach AS l JOIN CuDTx7 AS r ON l.xiCuItem=r.giCuItem
                                 WHERE 1=1 and NFileName like '%jpg' and bList = 'Y' and giCuItem = @xItem order by listSeq"; 

            using (var reader = SqlHelper.ReturnReader("ODBCDSN", sqlString3
                                     , DbProviderFactories.CreateParameter("ODBCDSN", "@xItem", "@xItem", xItem)))
            {

                if (reader.HasRows == false)
                {
                    this.Visible = false;
                    totalImagetItem.Add(new imageItem());
                }                 
                    while (reader.Read())
                    {
                        imageItem theitem = new imageItem();
                        theitem.image = "/public/Attachment/" + reader["NFileName"].ToString();
                        theitem.title = reader["aTitle"].ToString();
                        theitem.url = "";

                        totalImagetItem.Add(theitem);


                        if (reader["bgMusic"].ToString() == "Random")
                        {
                            if (MusicRandom.Count > 0)
                            {                                
                                Random rd = new Random();
                                getmusic = MusicRandom[rd.Next(0, MusicRandom.Count)];
                            }
                            else
                            {
                                //後台未設背景音樂
                                getmusic = string.Empty;
                            }
                        }
                        else
                        {
                           getmusic = reader["bgMusic"].ToString();                           
                        }
                    }                         

                result = Newtonsoft.Json.JsonConvert.SerializeObject(totalImagetItem);
            }
        }

    }

    private class imageItem
    {
        public string image { get; set; }
        public string title { get; set; }
        public string url { get; set; }
    }


}
