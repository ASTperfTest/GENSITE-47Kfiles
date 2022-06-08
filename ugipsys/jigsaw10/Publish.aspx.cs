using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;

public partial class Publish : System.Web.UI.Page
{
    public string returnUrl;
    protected void Page_Load(object sender, EventArgs e)
    {
        // LINQ to SQL (Repository)
        IRepository _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        // ReturnURL (Save&Return)
        returnUrl = getReturnURL();

        int id = int.Parse(Request["id"] ?? "0");

        if (IsPostBack)
        {
            string orderSiteUnit = Request.Form["orderSiteUnit"]?? "0";
            string orderSubject = Request.Form["orderSubject"]?? "0";
            string orderKnowledgeTank = Request.Form["orderKnowledgeTank"]?? "0";
            string orderKnowledgeHome = Request.Form["orderKnowledgeHome"]?? "0";

            #region //PostBack Data Check
            if (!CheckOrder(orderSiteUnit) || !CheckOrder(orderSubject) || !CheckOrder(orderKnowledgeTank) || !CheckOrder(orderKnowledgeHome))
            {
                StrFunc.GenErrMsg("議題關聯知識文章單元順序設定 => 輸入不正確！", string.Format("Javascript:window.history.back();"));
            }
            #endregion
            var result = from p in _mGIPcoanew_repository.List<KnowledgeJigsaw>().Where(p => p.parentIcuitem == id)
                         join s in _mGIPcoanew_repository.List<CuDTGeneric>() on p.gicuitem equals s.iCUItem
                         select new
                         {
                             s.iCUItem,
                             s.topCat,
                         };
            try
            {
                #region //UPDATE
                using (TransactionScope ts = new TransactionScope())
                {
                    foreach (var x in result)
                    {
                        if (Request.Form[x.iCUItem.ToString()] != null)
                        {
                            _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == x.iCUItem).fCTUPublic = Request.Form[x.iCUItem.ToString()][0];
                        }
                        if (x.topCat == "C")
                        {
                            var o = _mGIPcoanew_repository.Get<KnowledgeJigsaw>(p => p.gicuitem == x.iCUItem);
                            o.orderSiteUnit = int.Parse(orderSiteUnit);
                            o.orderSubject = int.Parse(orderSubject);
                            o.orderKnowledgeTank = int.Parse(orderKnowledgeTank);
                            o.orderKnowledgeHome = int.Parse(orderKnowledgeHome);
                        }
                    }
                    _mGIPcoanew_repository.Save();
                    ts.Complete();
                }
                #endregion
                Redirect();
            }
            catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }
        }
        else
        {
            #region //sql
            //sql="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("iCUItem")&"'"
            //sql1="SELECT KnowledgeJigsaw.gicuitem as gicuitem, CuDTGeneric.sTitle as sTitle, CuDTGeneric.fCTUPublic as fCTUPublic, KnowledgeJigsaw.orderSiteUnit as orderSiteUnit, KnowledgeJigsaw.orderSubject as orderSubject, KnowledgeJigsaw.orderKnowledgeTank as orderKnowledgeTank, KnowledgeJigsaw.orderKnowledgeHome as orderKnowledgeHome, KnowledgeJigsaw.parentIcuitem as parentIcuitem, KnowledgeJigsaw.gicuitem as gicuitem
            //FROM KnowledgeJigsaw
            //INNER JOIN CuDTGeneric ON KnowledgeJigsaw.gicuitem = CuDTGeneric.iCUItem
            //where KnowledgeJigsaw.parentIcuitem='"&request("iCUItem")&"'"
            #endregion
            var originCuDTGeneric = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);
            if (originCuDTGeneric != null)
            {
                titles.Text = string.Format("【內容區塊--{0}】", originCuDTGeneric.sTitle);

                var result = from p in _mGIPcoanew_repository.List<KnowledgeJigsaw>().Where(p => p.parentIcuitem == id)
                             join s in _mGIPcoanew_repository.List<CuDTGeneric>() on p.gicuitem equals s.iCUItem
                             select new
                             {
                                 contextOrder = s.topCat == "C" ? "<td colspan='3' class='eTableContent'>" + s.sTitle + "：&nbsp;入口網(站內單元)<input name='orderSiteUnit' value='" + (p.orderSiteUnit ?? 0) + "' maxlength='1' size='1'>&nbsp;主題館：<input name='orderSubject' value='" + (p.orderSubject ?? 0) + "' maxlength='1' size='1'>&nbsp;知識庫：<input name='orderKnowledgeTank' value='" + (p.orderKnowledgeTank ?? 0) + "' maxlength='1' size='1'>&nbsp;知識家：<input name='orderKnowledgeHome' value='" + (p.orderKnowledgeHome ?? 0) + "' maxlength='1' size='1'></td>" : "",
                                 contextTitle = s.topCat == "C" ? "" : "<td class='eTableContent'>" + s.sTitle + "</td>",
                                 contextPublic = s.topCat == "C" ? "" : "<td class='eTableContent'><select name='" + p.gicuitem + "' id='" + p.gicuitem + "'><option value='Y'" + (s.fCTUPublic == 'Y' ? " selected='selected'" : "") + ">公開</option><option value='N'" + (s.fCTUPublic == 'N' ? " selected='selected'" : "") + ">不公開</option></select></td>",
                                 contextAction = s.topCat == "A" ? "<td class='eTableContent'><input type='button' class='cbutton' onclick=\"location.href='PublishEdit.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='內容管理'>&nbsp;<input type='button' class='cbutton' onclick=\"location.href='PublishQuery.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='新增內容'></td>" :
                                                 s.topCat == "E" ? "<td class='eTableContent'><input type='button' class='cbutton' onclick=\"location.href='PublishEdit.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='內容管理'></td>" :
                                                 s.topCat == "F" ? "<td class='eTableContent'><input type='button' class='cbutton' onclick=\"location.href='PublishEdit.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='留言管理'></td>" :
                                                 s.topCat != "C" ? "<td class='eTableContent'><input type='button' class='cbutton' onclick=\"location.href='PublishEdit.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='內容管理'>&nbsp;<input type='button' class='cbutton' onclick=\"location.href='PublishQuery.aspx?pid=" + p.parentIcuitem + "&id=" + p.gicuitem + "&topCat=" + s.topCat + "'\" value='新增內容'></td>" :
                                                 "",
                             };
                
                rptList.DataSource = result;
                rptList.DataBind();
            }
            else
            {
                Redirect();
            }
        }
    }
    private string getReturnURL()
    {
        var result = Request.QueryString["returnUrl"] ?? "";
        if (result != "")
            Session.Add("returnUrl", result);
        else
            result = (Session["returnUrl"] ?? "Index.aspx").ToString();
        
        return result;
    }
    private void Redirect()
    {
        Response.Redirect(returnUrl);
    }
    private bool CheckOrder(string x)
    {
        return ("1234".Contains((x.Length > 0 ? x[0] : '0')));
    }
}
