using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Web.Configuration;
using GSS.Vitals.COA.Data;
using System.Data;

public class PhoneticNotation
{
    static PhoneticNotation()
    {
        string sql = @"
            if not exists(SELECT * FROM PhoneticNotation)
            begin 
	            insert into PhoneticNotation ([word], [PhoneticNotation] )
	            select N'鱒', N'ㄗㄨㄣ' union 
	            select N'鱲', N'ㄌㄚˋ' union 
	            select N'鯛', N'ㄉㄧㄠ' union 
	            select N'鯮', N'ㄗㄨㄥ' union 
	            select N'鱠', N'ㄏㄨㄟˋ' union 
	            select N'頜', N'ㄍㄜˊ' union 
	            select N'鰺', N'ㄙㄠ' union 
	            select N'鯔', N'ㄗ' union 
	            select N'鯧', N'ㄔㄤ' union 
	            select N'鮻', N'ㄙㄨㄛ' union 
	            select N'鰶', N'ㄐㄧˋ' union 
	            select N'棘', N'ㄐㄧˊ' union 
	            select N'蚶', N'ㄏㄢ' union 
	            select N'鰹', N'ㄐㄧㄢ' union 
	            select N'鮫', N'ㄐㄧㄠ' union 
	            select N'鰆', N'ㄔㄨㄣ' union 
	            select N'鰭', N'ㄑㄧˊ' union 
	            select N'鱝', N'ㄈㄣˋ' union 
	            select N'鯖', N'ㄑㄧㄥ' union 
	            select N'鮪', N'ㄨㄟˇ' union 
	            select N'魨', N'ㄊㄨㄣˊ' union 
	            select N'牡', N'ㄇㄨˇ' union 
	            select N'蛤', N'ㄍㄜˊ' union 
	            select N'蜆', N'ㄒㄧㄢˇ' union 
	            select N'鯡', N'ㄈㄟˋ' union 
	            select N'魬', N'ㄈㄢˇ' union 
	            select N'鯁', N'ㄍㄥˇ' union 
	            select N'鯙', N'ㄔㄨㄣˊ' 
            end";

        SqlHelper.ExecuteNonQuery("ODBCDSN", sql);
    }


    public static string GetNotation(string str)
    {
        string sqlString = @"SELECT [word], [PhoneticNotation] FROM PhoneticNotation";
        using (var result = SqlHelper.ReturnReader("ODBCDSN", sqlString))
        {
            while (result.Read())
            {
                str = str.Replace(result["word"].ToString(), result["word"].ToString() + "(" + result["PhoneticNotation"].ToString() + ")");
            }
        };
        return str;
    }
}