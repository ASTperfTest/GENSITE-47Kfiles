using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;

public partial class kmactivity_history_deletepic : System.Web.UI.Page
{
    ArrayList DirList = new ArrayList();
    protected void Page_Load(object sender, EventArgs e)
    {
        SearchDir("c:\\pic\\");
        SearchAllFile();
        Response.Write("over!!");
    }
    private void SearchDir(string path)
    {
        string[] Dirs = System.IO.Directory.GetDirectories(path);
        DirList.Add(path);
        for (int i = 0; i < Dirs.Length; i++)
        {
            SearchDir(Dirs[i]);
        }
    }
    ArrayList FileList = new ArrayList();

    void SearchAllFile()
    {
        for (int i = 0; i < DirList.Count; i++)
        {
            string[] Files = System.IO.Directory.GetFiles(DirList[i].ToString(), "*.*");

            for (int j = 0; j < Files.Length; j++)
            {
                if (Files[j].IndexOf("Thumbs.db") < 0)
                try
                {
                    System.Drawing.Image image = System.Drawing.Image.FromFile(Files[j]);
                    if (image.Height > image.Width)
                    {
                        image.Dispose();
                        System.IO.File.Delete(Files[j]);
                    }
                    else
                    {
                        image.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    Response.Write(Files[j]);
                }
            }
        }
    }
}