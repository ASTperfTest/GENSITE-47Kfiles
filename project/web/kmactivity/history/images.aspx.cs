using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing.Imaging;
using System.Drawing;

public partial class kmactivity_history_images : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || Session["memID"].ToString() =="")
        {
            Response.StatusCode = 403;
            Response.End();
        }
        Response.ContentType = "image/jpeg";
        Response.AppendHeader("Content-Disposition", "attachment; filename=test.jpg");
        DrowPicture();
    }

    private void DrowPicture()
    {
        string memid = Session["memID"].ToString().Trim();
        HistoryPicture historyPicture = new HistoryPicture(memid);
        HistoryPicture.UserQuestionNow huqn = historyPicture.GetUserQuestionNow();
        string imagepath = historyPicture.GetUserCurrentPic(huqn.Icuitem,huqn.CurrentNode);

        //System.Drawing.Image watermarkImage = System.Drawing.Image.FromFile(Server.MapPath("/images/present_02.jpg"));
        //作為浮水印的圖檔
        if (imagepath == "")
            Response.End();
        System.Drawing.Image image = System.Drawing.Image.FromFile(Server.MapPath("/public/History/" + imagepath));
        SolidBrush  mySolidBrush  = new SolidBrush (Color.Black);
        ImageFormat thisFormat = image.RawFormat;
        int fixWidth = 444;
        int fixHeight = 4444;
        int maxPx = 444;
        if (image.Width > maxPx || image.Height > maxPx)
        {
            if (image.Width >= image.Height)
            {
                fixWidth = maxPx;
                fixHeight = Convert.ToInt32((Convert.ToDouble(fixWidth) / Convert.ToDouble(image.Width)) * Convert.ToDouble(image.Height));
            }
            else
            {
                fixHeight = maxPx;
                fixWidth = Convert.ToInt32((Convert.ToDouble(fixHeight) / Convert.ToDouble(image.Height)) * Convert.ToDouble(image.Width));
            }
        }
        else
        {
            if (image.Height > image.Width)
            {
                fixHeight = image.Height + (image.Height % 3);
                fixWidth = image.Width + (image.Width % 2); ;
            }
            else
            {
                fixHeight = image.Height + (image.Height % 2);
                fixWidth = image.Width + (image.Width % 3); ;
            }
        }

        Bitmap imageOutput = new Bitmap(image, fixWidth, fixHeight);
        Graphics gra = Graphics.FromImage(imageOutput);
        //宣告出一個GDI
        //gra.DrawImage(watermarkImage, new Rectangle(imageOutput.Width - watermarkImage.Width, imageOutput.Height - watermarkImage.Height, imageOutput.Width, imageOutput.Height), 0, 0, imageOutput.Width, imageOutput.Height, GraphicsUnit.Pixel);
        //重新繪製被縮的圖並加浮水印上去
       // gra.FillRectangle(mySolidBrush, 0, 0, 100, 100);
        if (huqn.STATES != 1 && huqn.STATES != 3)
        {
            Rectangle[] rects = new Rectangle[6];
            PictureType v = new PictureType();
            int i = 0;
            int imgw = fixWidth / 3;
            int imgh = fixHeight / 2;
            if (fixHeight > fixWidth)
            {
                imgw = fixWidth / 2;
                imgh = fixHeight / 3;
            }
            foreach (int value in Enum.GetValues(typeof(PictureType)))
            {
                if ((value & huqn.Picturemap) != value)
                {
                    int t = i / 3;
                    int y = i - (t * 3);
                    if (fixHeight > fixWidth)
                    {
                        rects[i] = new Rectangle(0 + imgw * t, 0 + imgh * y, imgw, imgh);
                    }
                    else
                    {
                        rects[i] = new Rectangle(0 + imgw * y, 0 + imgh * t, imgw, imgh);
                    }
                }
                i++;
            }
            gra.FillRectangles(mySolidBrush, rects);
            string fixSaveName = string.Concat("harvesthistory", ".jpg");
        }
        EncoderParameter para = new EncoderParameter(Encoder.Quality, 100L);
        EncoderParameters paras = new EncoderParameters(1);
        paras.Param[0] = para;
        ImageCodecInfo code = GetEncoderInfo("image/jpeg");
        imageOutput.Save(Response.OutputStream, code, paras);
        imageOutput.Dispose();
        image.Dispose();
        gra.Dispose();

    }
    private static ImageCodecInfo GetEncoderInfo(String mimeType)
    {
        int j;
        ImageCodecInfo[] encoders;
        encoders = ImageCodecInfo.GetImageEncoders();
        for (j = 0; j < encoders.Length; ++j)
        {
            if (encoders[j].MimeType == mimeType)
                return encoders[j];
        }
        return null;
    }

    public enum PictureType 
    {
        ONE = 1,
        TWO = 2,
        THREE = 4,
        FOUR = 8,
        FIVE = 16,
        SIX = 32
    }
}