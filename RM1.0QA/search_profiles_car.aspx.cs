using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections.Specialized;

public partial class search_profiles_car : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Create an ASPSessionVar object,
        //passing in the current context
        /*ASPSessionVar oASPSessionVar = new ASPSessionVar(Context);
        string sTemp = oASPSessionVar.GetSessionVar("pro_con");
        HttpCookie cookie = Request.Cookies["rate%2Dmonitor%2Ecom"];
        if (cookie != null)
        {
            if (Request.Cookies["rate%2Dmonitor%2Ecom"]["live%5Fsession"] != "auto" || sTemp == "")
                Response.Redirect("default_session.asp");
        }
        else
            Response.Redirect("default_session.asp");

        //Grid1.SelectCommand += new ComponentArt.Web.UI.Grid.GridItemEventHandler(Grid1_SelectCommand);
        Grid1.SortCommand += new ComponentArt.Web.UI.Grid.SortCommandEventHandler(Grid1_SortCommand);
        Grid1.FilterCommand += new ComponentArt.Web.UI.Grid.FilterCommandEventHandler(Grid1_FilterCommand);
        CarShopRequestSource.Selecting += new SqlDataSourceSelectingEventHandler(CarShopRequestSource_Selecting);
         */
    }

    void CarShopRequestSource_Selecting(object sender, SqlDataSourceSelectingEventArgs e)
    {
        e.Command.Parameters["@user_id"].Value = Request.Cookies["rate%2Dmonitor%2Ecom"]["user%5Fid"];
        //Response.Write(e.Command.Parameters["@user_id"].Value.ToString());
    }

    public string UserName
    {
        get
        {
            return Request.Cookies["rate%2Dmonitor%2Ecom"]["user%5Fname"].Replace('+', ' ');
        }
    }

    void Grid1_FilterCommand(object sender, ComponentArt.Web.UI.GridFilterCommandEventArgs e)
    {
        string expression = e.FilterExpression;

        foreach (ComponentArt.Web.UI.GridColumn column in Grid1.Levels[0].Columns)
        {
            if (column.IsSearchable)
                expression = expression.Replace(column.DataField, "CONVERT(" + column.DataField + ", 'System.String')");
        }

        Grid1.Filter = expression;

    }

    void Grid1_SortCommand(object sender, ComponentArt.Web.UI.GridSortCommandEventArgs e)
    {
        Grid1.Sort = e.SortExpression;
    }

    private void Grid1_DoubleClick(object sender, ComponentArt.Web.UI.GridItemEventArgs args)
    {
        Response.Write("<script language='javascript'>window.open('alerts_rate_management_car.asp');</script>");
    }
}
