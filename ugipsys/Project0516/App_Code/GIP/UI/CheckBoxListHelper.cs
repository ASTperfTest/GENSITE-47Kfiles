using System;
using System.Collections;

/// <summary>
/// CheckBoxListHelper 的摘要描述
/// </summary>
public class CheckBoxListHelper
{
	public CheckBoxListHelper()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}

	public static void setSelected(System.Web.UI.WebControls.CheckBoxList checkBoxList, IList values)
	{
		if(values == null)
			return;

		foreach (System.Web.UI.WebControls.ListItem item in checkBoxList.Items)
		{
			if (values.Contains(item.Value))
				item.Selected = true;
		}
	}

	public static IList getSelectedValues(System.Web.UI.WebControls.CheckBoxList checkBoxList)
	{
		IList results = new ArrayList();

		foreach (System.Web.UI.WebControls.ListItem item in checkBoxList.Items)
		{
			if (item.Selected)
				results.Add(item.Value);
		}

		return results;
	}
}
