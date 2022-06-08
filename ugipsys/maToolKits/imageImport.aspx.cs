using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class imageImport : System.Web.UI.Page
{
    // 農業百年發展史 ctRootId 
    private string ctRootId = System.Web.Configuration.WebConfigurationManager.AppSettings["ctRootID"].ToString();
    // 珍貴老照片 pic_UnitId & pic_NodeId
    private string pic_UnitId = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
    private string pic_NodeId = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTNodeId"].ToString();
    // UserName
    private string MemberID = "hyweb";
    private string Script = "<script>alert('連線逾時或尚未登入，請重新登入');window.close();</script>";
    // DataLevel
    private int DataLevel = 1;
    // Depth
    private int Depth = 0;
    // DataParent 預設為 珍貴老照片
    private string DataParent = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTNodeId"].ToString();
    // tempTargetDirectory
    private string tempTargetDir;
    // 程式位置
    private string programDir = "ugipsys\\matoolkits";
    // 照片位置
    private string TargetDir = System.Web.Configuration.WebConfigurationManager.AppSettings["TargetDir"].ToString();

    protected void Page_Load(object sender, EventArgs e)
    {
        //if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
        //{
        //    Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
        //    return;
        //}
        //else
        //{
        //    MemberID = Session["userID"].ToString();
        //}
    }

    protected void btnImport_Click(object sender, EventArgs e)
    {
        Response.Buffer = false;
        try
        {
            // ../ugipsys/maToolKits/imageImport (欲匯入資料夾路徑)
            string activeDir = Server.MapPath("imageImport/");
            TargetDir = Server.MapPath("../" + TargetDir);
            //Response.Write(TargetDir);
            ProcessDirectory(activeDir);

            Response.Write("匯入完成");
        }
        catch (Exception)
        {
            throw;
        }
    }

    // 改名
    private string dirReName(string OldName)
    {
        // do Directory ReName..
        return null;
    }

    // 巡覽資料夾(Recursion)
    private void ProcessDirectory(string dirFullPath)
    {

        //Response.Write(subDir.Length+ "<br/>");
        //Response.Write(dirFullPath);
        //Response.Write("   " + Depth + "<br/>");
        DirectoryInfo dirInfo = new DirectoryInfo(dirFullPath);
        var subDir = Directory.GetDirectories(dirFullPath);

        if (dirInfo.Name != "imageImport")
        {
            // 建立資料夾
            //Directory.CreateDirectory(tempTargetDir + dirInfo.Name);
            // 更新最新的目標資料夾
            tempTargetDir += dirInfo.Name + "\\";
            

            //Response.Write(tempTargetDir + "<br/>");

            if (subDir.Length != 0) // 珍貴老照片的子節點
            {
                //DataParent = CreateDir2Database(dirReName(dirInfo.Name),  DataLevel+ Depth, pic_NodeId);
                DataParent = CreateDir2Database(dirInfo.Name, DataLevel + Depth, DataParent, "C");
            }
            else // 珍貴老照片的子節點的子節點
            {
                //DataParent = CreateDir2Database(dirReName(dirInfo.Name), DataLevel + Depth, DataParent);
                DataParent = CreateDir2Database(dirInfo.Name, DataLevel + Depth, DataParent, "U");
            }

            ProcessDirFiles(dirFullPath, DataParent);
        }
        foreach (string path in subDir)
        {
            Depth++;
            ProcessDirectory(path);
        }
        Depth--;

        // 設定DataParent
        if (dirInfo.Parent.ToString() != "imageImport")
        {
            string strQueryDataParent = @"SELECT DataParent FROM CatTreeNode WHERE CtNodeId = @CtNodeId";
            DataParent = SqlHelper.ReturnScalar("ConnString", strQueryDataParent,
                DbProviderFactories.CreateParameter("ConnString", "@CtNodeId", "@CtNodeId", DataParent)).ToString();
        }
        else
        {
            DataParent = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTNodeId"].ToString();
        }
    }

    // 建立資料夾
    private string CreateDir2Database(string dirName, int Depth, string dtParent, string NodeType)
    {
        ScrapeASCII(ref dirName);
        string strInsertQuery = string.Empty;
        if (NodeType != "U")
        {
            strInsertQuery = @"INSERT INTO CattreeNode (CtRootID, CatName, DataLevel, DataParent, CtUnitID, EditUserID)
                                    VALUES (@CtRootID, @CatName, @DataLevel, @DataParent, @CtUnitID, @EditUserID) 
                                  SELECT SCOPE_IDENTITY()";
        }
        else
        {
            strInsertQuery = @"INSERT INTO CattreeNode (CtRootID,CtNodeKind, CatName, DataLevel, DataParent, CtUnitID, EditUserID)
                                    VALUES (@CtRootID, @CtNodeKind, @CatName, @DataLevel, @DataParent, @CtUnitID, @EditUserID) 
                                  SELECT SCOPE_IDENTITY()";
        }

        
        var newCtNodeId = SqlHelper.ReturnScalar("ConnString", strInsertQuery,
              DbProviderFactories.CreateParameter("ConnString", "CtRootID", "CtRootID", ctRootId),
              DbProviderFactories.CreateParameter("ConnString", "CtNodeKind", "CtNodeKind", NodeType),
              DbProviderFactories.CreateParameter("ConnString", "CatName", "CatName", dirName),
              DbProviderFactories.CreateParameter("ConnString", "DataLevel", "DataLevel", Depth),
              DbProviderFactories.CreateParameter("ConnString", "DataParent", "DataParent", dtParent),
              DbProviderFactories.CreateParameter("ConnString", "CtUnitID", "CtUnitID", pic_UnitId),
              DbProviderFactories.CreateParameter("ConnString", "EditUserID", "EditUserID", MemberID));

        return newCtNodeId.ToString();
    }

    int serialNo = 0;

    // 建立檔案
    private void ProcessDirFiles(string dirPath, string refID)
    {
        //import files to database....
        string iBaseDSD = "7";
        FileInfo[] files = new DirectoryInfo(dirPath).GetFiles();
        foreach (FileInfo file in files)
        {
            if (file.Name != "Thumbs.db")
            {
                serialNo++;
                string newName = "H" + DateTime.Now.Millisecond.ToString() + serialNo.ToString();
                file.CopyTo(TargetDir + newName + file.Extension);
                string strInsertScript = @"INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, sTitle, iEditor, 
                                            refID, xImgFile, iDept, fCTUPublic, xBody)
                                       VALUES (@iBaseDSD, @iCTUnit, @sTitle, @iEditor, 
                                            @refID, @xImgFile, @iDept, @fCTUPublic, @xBody)
                                           SELECT SCOPE_IDENTITY()";
               SqlHelper.ExecuteNonQuery("ConnString", strInsertScript,
                    DbProviderFactories.CreateParameter("ConnString", "@iBaseDSD", "@iBaseDSD", iBaseDSD),
                    DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", pic_UnitId),
                    DbProviderFactories.CreateParameter("ConnString", "@sTitle", "@sTitle", newName),
                    DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID),
                    DbProviderFactories.CreateParameter("ConnString", "@refID", "@refID", refID),
                    DbProviderFactories.CreateParameter("ConnString", "@xImgFile", "@xImgFile", newName + file.Extension),
                    DbProviderFactories.CreateParameter("ConnString", "@iDept", "@iDept", "0"),
                    DbProviderFactories.CreateParameter("ConnString", "@fCTUPublic", "@fCTUPublic", "Y"),
                    DbProviderFactories.CreateParameter("ConnString", "@xBody", "@xBody", file.FullName)
                    );

               Response.Write(file.FullName + "<br/>");
            }
        }
    }

    // 刪除全部資料
    protected void btnClearDB_Click(object sneder, EventArgs e)
    {
        // 先取得珍貴老照片底下的節點 (不含珍貴老照片)
        string strRecursion = @"with node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) as
                                (
	                                SELECT ctNodeId,dataParent,CatName,DataLevel,cast('' as varchar(max))
	                                FROM CatTreeNode
	                                WHERE ctNodeId = @currentNodeId AND ctRootID = @ctRootId
	                                UNION ALL
	                                SELECT CatTreeNode.ctNodeId,node_Tree.ctNodeId
                                        ,CatTreeNode.CatName,CatTreeNode.DataLevel
		                                ,cast(node_Tree.fullPath + '/' + CatTreeNode.CatName as varchar(max))
	                                FROM CatTreeNode
	                                INNER JOIN node_Tree ON CatTreeNode.dataParent = node_Tree.ctNodeId
                                )
                                SELECT * FROM node_Tree WHERE ctNodeId <> @currentNodeId";

        using (var reader = SqlHelper.ReturnReader("ConnString", strRecursion,
            DbProviderFactories.CreateParameter("ConnString", "@currentNodeId", "@currentNodeId", pic_NodeId),
            DbProviderFactories.CreateParameter("ConnString", "@ctRootId", "@ctRootId", ctRootId)))
        {
            while (reader.Read())
            {
                string strDeleteRecord = @"DELETE FROM CuDTGeneric WHERE iCTUnit = @iCTUnit AND refID = @refID";
                string strDeleteNode = @"DELETE FROM CatTreeNode WHERE CtNodeID = @CtNodeID";
                // 刪除節點底下文章
                SqlHelper.ExecuteNonQuery("ConnString", strDeleteRecord,
                    DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", pic_UnitId),
                    DbProviderFactories.CreateParameter("ConnString", "@refID", "@refID", reader["ctNodeId"].ToString()));
                // 刪除該節點
                SqlHelper.ExecuteNonQuery("ConnString", strDeleteNode,
                    DbProviderFactories.CreateParameter("ConnString", "@CtNodeID", "@CtNodeID", reader["ctNodeId"].ToString()));
            }
        }
        Response.Write("刪除完成");
    }


    private void ScrapeASCII(ref string str)
    {
	
		if (str == "1975豐年兔展@榮星花園"  ||
		    str == "Reynolds（藍瑙）女士參觀農村廚房" ||
		    str == "Billings（畢玲絲）小姐實施農村調查" ||
		    str == "1958" ||
		    str == "brown與brickling1961")
		{
			return;
		}
	
	
        string r = string.Empty;
        foreach (var c in str)
        {
            int ascii = c;
            if (ascii > 255)
            {
                r += c;
            }
        }

        str = r;
    }

    protected void ProcessWhiteList(object sender, EventArgs e)
    {

        SqlHelper.ExecuteNonQuery("ConnString", "delete from HISTORY_PICTURE..picture_node");

        string str = @"
        insert into  HISTORY_PICTURE..picture_node (nodeid)
        select top 1 refid from mGIPcoanew..CuDTGeneric 
        where ictunit = @iCTUnit
        and xbody like ('%' + @title + '%')
        and not refid in (select nodeid from HISTORY_PICTURE..picture_node)";
        
        Response.Buffer = false;
        foreach (var item in whiteList)
        {
            SqlHelper.ExecuteNonQuery("ConnString", str
                , DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", pic_UnitId)
                , DbProviderFactories.CreateParameter("ConnString", "@title", "@title", item));
                       
        }
    }

    private List<string> whiteList = new List<string>(){
         @"01 糧食作物\01-02.甘藷"
        ,@"02 糖用作物\03-製糖、加工、試驗"
        ,@"03 嗜好作物\03-01 茶"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01水土保持"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01收穫運送"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01香蕉繩索"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01套袋"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01病蟲害"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01植株"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01噴藥"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01蕉園"
        ,@"09 常綠作物\09-01 香蕉\09-01-01香蕉植栽\09-01灌溉"
        ,@"09 常綠作物\09-01 香蕉\09-01-02蕉園防颱\09-01-02支柱處理"
        ,@"09 常綠作物\09-01 香蕉\09-01-02蕉園防颱\09-01-02未立支柱災後"
        ,@"09 常綠作物\09-01 香蕉\09-01-02蕉園防颱\09-01-02立支柱"
        ,@"09 常綠作物\09-01 香蕉\09-01-03香蕉果指"
        ,@"09 常綠作物\09-01 香蕉\09-01-05香蕉運銷\香蕉外銷"
        ,@"09-01 香蕉\09-01-05香蕉運銷\香蕉運銷"
        ,@"09 常綠作物\09-03 鳳梨\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-02日曬防治"
        ,@"09 常綠作物\09-03 鳳梨\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-03水土保持"
        ,@"09 常綠作物\09-03 鳳梨\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-04田園風光及山坡地平地果園"
        ,@"09 常綠作物\09-03 鳳梨\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-05田園觀摩及人物"
        ,@"09 常綠作物\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-11種植"
        ,@"09 常綠作物\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-12鳳梨性狀"
        ,@"09 常綠作物\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-13鳳梨芽"
        ,@"09 常綠作物\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-14澆水電石水催芽"
        ,@"09 常綠作物\09-03 鳳梨\09-03-01-鳳梨植育\09-03-01-15整地"
        ,@"09 常綠作物\09-03 鳳梨\09-03-02-鳳梨採收運銷\09-03-02-01市集"
        ,@"09 常綠作物\09-03 鳳梨\09-03-02-鳳梨採收運銷\09-03-02-03田間半機械採收"
        ,@"09 常綠作物\09-03 鳳梨\09-03-02-鳳梨採收運銷\09-03-02-05集貨"
        ,@"09 常綠作物\09-03 鳳梨\09-03-03鳳梨加工\09-03-03-01鳳梨工廠"
        ,@"09 常綠作物\09-04-荔枝\09-04-03市集及產品"
        ,@"09 常綠作物\09-05-龍眼\09-05-02-共同運銷市集分級"
        ,@"11 蔬菜\11-10 果菜類\11-10-02 西瓜"
        ,@"16 樹木\16-04 造林、林場"
        ,@"16 樹木\16-05-07樟木"
        ,@"18 畜牧\18-01 牛\04 其他牛的照片"
        ,@"18 畜牧\18-03 豬\18-03-03 豬隻"
        ,@"18 畜牧\18-03 豬\18-03-06 豬隻運銷"
        ,@"18 畜牧\18-04.兔\美國軍機載送紐西蘭兔抵台"
        ,@"18 畜牧\18-19.驢"
        ,@"20 飼料\豆餅"
        ,@"23 水產\23-04 漁船、漁具"
        ,@"23 水產\23-05 魚市、運銷"
        ,@"23 水產\23-06 魚苗飼育"
        ,@"23 水產\23-08 食用魚\23-08-11 烏魚、烏魚子"
        ,@"23 水產\23-09 貝介類\23-09-01 採牡蠣"
        ,@"29 植物保護\29-10 空中噴藥施肥\水稻紋枯病防治"
        ,@"29 植物保護\29-10 空中噴藥施肥\其他空中噴藥勁照"
        ,@"29 植物保護\29-10 空中噴藥施肥\農業航空公司領航過程"
        ,@"30 肥料\30-02 化學肥料\美援肥料"
        ,@"31 加工\31-01 奶品加工\乳品製程"
        ,@"31-03 皮蛋\湖內鄉農會鴨蛋加工廠-皮蛋製程"
        ,@"31 加工\31-07 玉米加工\玉米筍及製程"
        ,@"31 加工\31-07 玉米加工\玉米罐頭及製程"
        ,@"31 加工\31-09 豆類加工\大豆\31-09豆腐及製程"
        ,@"31 加工\31-09 豆類加工\大豆\31-09豆瓣醬及製程"
        ,@"31 加工\31-09 豆類加工\大豆\31-09醬油及製程"
        ,@"31 加工\31-09 豆類加工\綠豆製品"
        ,@"31 加工\31-10 蔬菜加工\菇類加工"
        ,@"31 加工\31-10 蔬菜加工\豌豆加工"
        ,@"31 加工\31-10 蔬菜加工\蘆筍加工"
        ,@"31 加工\31-10 蔬菜加工\蘿蔔加工"
        ,@"31 加工\31-12 其他加工\晒豆皮"
        ,@"32 機具\32-01 農用機具\32-01-02 耕耘機\女性使用耕耘機"
        ,@"32 機具\32-01 農用機具\32-01-03 插秧機\女性使用插秧機"
        ,@"32 機具\32-01 農用機具\32-01-21 牛車"
        ,@"32 機具\32-02 非農業專用機具\32-02-21 收音機\農復會補助6000台收音機"
        ,@"36.農家\36-01 農家房舍\菸樓"
        ,@"36.農家\36-01 農家房舍\農家房舍(補充)"
        ,@"36.農家\36-02 農家生活\市集買賣"
        ,@"36.農家\36-02 農家生活\老人娛樂"
        ,@"36.農家\36-02 農家生活\其他"
        ,@"36.農家\36-02 農家生活\居家生活"
        ,@"36.農家\36-02 農家生活\農村生活"
        ,@"36.農家\36-02 農家生活\農村生活(宗教祭祀)\文化體驗"
        ,@"36.農家\36-02 農家生活\農村生活(宗教祭祀)\炸寒單"
        ,@"36.農家\36-02 農家生活\農村生活(宗教祭祀)\捐助金身"
        ,@"36.農家\36-02 農家生活\農村生活(宗教祭祀)\過火(過香灰)"
        ,@"36.農家\36-02 農家生活\農村社會生活"
        ,@"36.農家\36-02 農家生活\農村清潔衛生教育"
        ,@"36.農家\36-02 農家生活\農業調查"
        ,@"36.農家\36-02 農家生活\精神生活(樂讀豐年)"
        ,@"36.農家\36-04 廚房設備\Reynolds（藍瑙）女士參觀農村廚房"
        ,@"36.農家\36-04 廚房設備\水槽演進史"
        ,@"36.農家\36-04 廚房設備\農家使用瓦斯"
        ,@"36.農家\36-06 保健衛生\口腔保健"
        ,@"36.農家\36-06 保健衛生\老人照護"
        ,@"36.農家\36-06 保健衛生\育嬰"
        ,@"36.農家\36-06 保健衛生\急救訓練"
        ,@"36.農家\36-06 保健衛生\急救教育"
        ,@"36.農家\36-06 保健衛生\捐血"
        ,@"36.農家\36-06 保健衛生\健康檢查"
        ,@"36.農家\36-06 保健衛生\預防接種"
        ,@"37.家政\37-05 烹飪（品嚐、烹飪、盤菜）\烹飪\烹飪-家政教育"
        ,@"37.家政\37-06 糕餅麵食\米食\狀元糕"
        ,@"37.家政\37-06 糕餅麵食\米食\飯糰"
        ,@"37.家政\37-06 糕餅麵食\零食、點心\豆花"
        ,@"37.家政\37-06 糕餅麵食\零食、點心\蛋捲"
        ,@"37.家政\37-06 糕餅麵食\零食、點心\爆米香"
        ,@"37.家政\37-06 糕餅麵食\麵食\油條"
        ,@"37.家政\37-06 糕餅麵食\麵食\潤餅皮"
        ,@"37.家政\37-09 木工、手工藝品、新產品\手工藝\竹籠製作"
        ,@"37.家政\37-09 木工、手工藝品、新產品\木工\吉他、小提琴製作"
        ,@"37.家政\37-09 木工、手工藝品、新產品\木工\馬桶"
        ,@"37.家政\37-09 木工、手工藝品、新產品\木工\桶"
        ,@"37.家政\37-10 玩偶（玩具）\玩偶\布袋戲"
        ,@"37.家政\37-10 玩偶（玩具）\玩偶\皮影戲"
        ,@"37.家政\37-13 護理\護理教學與實務"
        ,@"37.家政\37-20 家庭訪視\Billings（畢玲絲）小姐實施農村調查"
        ,@"38.四健會、農會\38-06 草根大使\「美國草地人」劇照"
        ,@"38.四健會、農會\38-06 草根大使\洗牛的草根大使"
    };
}