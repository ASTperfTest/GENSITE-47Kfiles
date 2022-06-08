using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using Jayrock.Json;
using Jayrock.Json.Conversion;
using Newtonsoft.Json.Linq;
using System.Text;
using GSS.Vitals.API.Client;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using System.Collections;

public partial class services_categoryservices : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        DispatchMethod("expandtreenode");
    }

    private void DispatchMethod(string MethodName)
    {
        switch (MethodName)
        {
            case "expandtreenode":
                ExpandOUTreeNode();
                break;
            default:
                break;
        }
    }

    private void ExpandOUTreeNode()
    {
        string CategoryId = WebUtility.GetStringParameter("nodeid");
        RestRpcClient apiClient = WebUtility.GetAPIClient();
        int parseInt=2;
        try
        {
            if (!string.IsNullOrEmpty(CategoryId))
            {
                if (!int.TryParse(CategoryId, out parseInt))
                {
                    Response.Redirect("/mp.asp?mp=1");
                    Response.End();
                }
            }
            CategoryInfoPagedCollection catrgoryinfo = apiClient.categories.GetChildren(parseInt, false);
            IList categoryList = new ArrayList();
            foreach (CategoryInfo cic in catrgoryinfo.Elements)
            {
                ExpertCategory ec = new ExpertCategory();
                ec.DisplayName = cic.DisplayName;
                ec.Id = cic.CategoryId;
                categoryList.Add(ec);
            }
            WebUtility.WriteAjaxResult(true, "", categoryList);
        }catch(Exception ex)
        {
            WebUtility.WriteAjaxResult(false, "", null);
        }
    }

    class ExpertCategory
    {
        string m_DisplayName;
        int m_Id;
        public string DisplayName
        {
            get { return m_DisplayName; }
            set { m_DisplayName = value; }
        }
        public int Id
        {
            get { return m_Id; }
            set { m_Id = value; }
        }
    }
}