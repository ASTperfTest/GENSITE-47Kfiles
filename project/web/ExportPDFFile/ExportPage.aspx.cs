using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web.UI;


public partial class ExportPDFFile_ExportPage : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
		
    }

	private DataTable GetFileDataTable()
	{
		string connStr = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;

		DataTable dtResult = new DataTable();

		using (SqlConnection conn = new SqlConnection(connStr))
		{
			string cmdText = @"SELECT CuDTGeneric.iCUItem,CuDTGeneric.sTitle,CuDTGeneric.fileDownLoad,dbo.CuDTAttach.NFileName 
							FROM CuDTGeneric
							INNER JOIN dbo.CuDTAttach ON dbo.CuDTGeneric.iCUItem = dbo.CuDTAttach.xiCuItem
							WHERE iCTUnit = 2766 
							AND (NFileName LIKE '%.pdf' OR NFileName LIKE '%.PDF')
							UNION ALL
							SELECT CuDTGeneric.iCUItem,CuDTGeneric.sTitle,CuDTGeneric.fileDownLoad,dbo.CuDTAttach.NFileName 
							FROM CuDTGeneric
							INNER JOIN dbo.CuDTAttach ON dbo.CuDTGeneric.iCUItem = dbo.CuDTAttach.xiCuItem
							WHERE iCTUnit = 135 
							AND (NFileName LIKE '%.pdf' OR NFileName LIKE '%.PDF')
							UNION ALL
							SELECT CuDTGeneric.iCUItem,CuDTGeneric.sTitle,CuDTGeneric.fileDownLoad,dbo.CuDTAttach.NFileName
							FROM CuDTGeneric
							INNER JOIN dbo.CuDTAttach ON dbo.CuDTGeneric.iCUItem = dbo.CuDTAttach.xiCuItem
							WHERE iCTUnit = 126 
							AND (NFileName LIKE '%.pdf' OR NFileName LIKE '%.PDF')";

			conn.Open();

			SqlCommand cmd = new SqlCommand(cmdText, conn);
			SqlDataAdapter ada = new SqlDataAdapter(cmdText, conn);
			ada.Fill(dtResult);

		}

		return dtResult;


		
 
	}
	protected void btnCalTotal_Click(object sender, EventArgs e)
	{
		DataTable dtCopedFile = GetFileDataTable();

		labelTotalFileCount.Text = "檔案筆數共：" + dtCopedFile.Rows.Count.ToString() +"筆";

	}
	protected void btnCopy_Click(object sender, EventArgs e)
	{
		DataTable dtCopedFile = GetFileDataTable();

		string fileName = "";
		string sourcePath = "";
		string targetPath = "";

		
		string serverPath = Server.MapPath("~");
		serverPath = serverPath.Substring(0, serverPath.LastIndexOf("web"));
		serverPath = serverPath + @"sys\public\Attachment";
		
		sourcePath = serverPath;
		targetPath = serverPath + "\\Copy";
		string logPath = serverPath + "\\Copy\\log.txt";
		int successCount = 0;
		int failCount = 0;

		if (!System.IO.Directory.Exists(targetPath))
		{
			System.IO.Directory.CreateDirectory(targetPath);
		}

		string notFoundFile = string.Empty;
		foreach (DataRow dr in dtCopedFile.Rows)
		{
			fileName = dr["NFileName"].ToString();

			string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
			string destFile = System.IO.Path.Combine(targetPath, fileName);

			if (File.Exists(sourceFile))
			{
				System.IO.File.Copy(sourceFile, destFile, true);
				successCount++;
			}
			else
			{
				notFoundFile = notFoundFile + (failCount+1).ToString() + ":" + dr["iCUItem"].ToString() + "_" + dr["NFileName"].ToString() + "\r\n";

				failCount++;
			}
		}

		File.WriteAllText(logPath, notFoundFile);
		
		//string alertMsg = "alert('成功複製：" + successCount + "筆，失敗：" + failCount + "筆')";
		string alertMsg = "成功複製：" + successCount + "筆，失敗：" + failCount + "筆";
		labelExportResult.Text = alertMsg;
		//ScriptManager.RegisterStartupScript(this.Page, GetType(), "ExportFileScript", alertMsg, true);

	}
	protected void Button1_Click(object sender, EventArgs e)
	{
		string serverPath = Server.MapPath("~");
		serverPath = serverPath.Substring(0, serverPath.LastIndexOf("web"));
		labelFilePath.Text  = serverPath + @"sys\public\Attachment";
		
	}
}
