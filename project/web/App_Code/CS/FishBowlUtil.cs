using System;
using System.IO;
using System.Collections.Generic;
using System.Web;
using System.Text;
using System.Security.Cryptography;
using System.Xml;
using System.Net;
using System.Text.RegularExpressions;

public class FishBowlUtil
{
	private FishBowlUtil(){}

    /**
     * 驗證是否符合編碼
     */
    public static bool isMatchedKey(string _login_id, string _game_key, string _key)
    {
        string _salt = "hello i am eddie";
        string _hashed_key = MD5(_login_id + _salt + _game_key);
        if (_key == _hashed_key)
        {
            return true;
        }
        return false;
    }

    /**
     * MD5編碼
     */
    public static string MD5(string Value)
    {
        MD5CryptoServiceProvider x = new MD5CryptoServiceProvider();
        byte[] data = Encoding.ASCII.GetBytes(Value);
        data = x.ComputeHash(data);
        string ret = "";
        for (int i = 0; i < data.Length; i++)
        {
            ret += data[i].ToString("x2").ToLower();
        }
        return ret;
    }

    /**
     * 取回時間差秒數
     */
    public static long getTimeDelta(DateTime pass_time)
    {
        TimeSpan t = DateTime.Now - pass_time;
        int game_speed = System.Convert.ToInt32(getConfig("game_speed"));
        return ((long)t.TotalSeconds * game_speed);
    }

    public static string getConfig(string tag_name)
    {
        WebClient wc = new WebClient();
        string config_file = System.Web.HttpContext.Current.Server.MapPath("assets/config.xml");

        wc.Encoding = UTF8Encoding.UTF8;
        string result = wc.DownloadString(config_file);

        XmlDocument doc = new XmlDocument();

        doc.LoadXml(result);

        return doc.GetElementsByTagName(tag_name)[0].InnerText;
    }

    public static int calcStars(int fish_type, int score)
    {
        int stars = 0;

        switch(fish_type)
        {
            // anemone
            case 1:
                if (score <= 130)
                {
                    stars = 1;
                }
                else if (score >= 131 && score <= 180)
                {
                    stars = 3;
                }
                else if (score >= 181 && score <= 220)
                {
                    stars = 5;
                }
                else if (score >= 221 && score <= 260)
                {
                    stars = 7;
                }
                else if (score >= 261)
                {
                    stars = 10;
                }
                break;

            // bream
            case 2:
                if (score <= 40)
                {
                    stars = 1;
                }
                else if (score >= 41 && score <= 90)
                {
                    stars = 3;
                }
                else if (score >= 91 && score <= 170)
                {
                    stars = 5;
                }
                else if (score >= 171 && score <= 210)
                {
                    stars = 7;
                }
                else if (score >= 211)
                {
                    stars = 10;
                }
                break;

            // clownfish
            case 3:
                if (score <= 180)
                {
                    stars = 1;
                }
                else if (score >= 181 && score <= 320)
                {
                    stars = 3;
                }
                else if (score >= 321 && score <= 420)
                {
                    stars = 5;
                }
                else if (score >= 421 && score <= 500)
                {
                    stars = 7;
                }
                else if (score >= 501)
                {
                    stars = 10;
                }
                break;

            // hippocampus
            case 4:
                if (score <= 280)
                {
                    stars = 1;
                }
                else if (score >= 281 && score <= 380)
                {
                    stars = 3;
                }
                else if (score >= 381 && score <= 460)
                {
                    stars = 5;
                }
                else if (score >= 461 && score <= 540)
                {
                    stars = 7;
                }
                else if (score >= 541)
                {
                    stars = 10;
                }
                break;
        }

        return stars;
    }

    public static string get_ip()
    {
        return System.Web.HttpContext.Current.Request.UserHostAddress;
    }
}
