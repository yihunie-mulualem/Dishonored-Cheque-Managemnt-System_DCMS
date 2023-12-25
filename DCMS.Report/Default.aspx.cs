using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;

using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Web;
using System.Runtime.Remoting.Contexts;
using System.Data.SqlClient;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {
        DateTime from;
        DateTime to;
        string fromVal = Request.QueryString["from"];
        string toVal = Request.QueryString["to"];
        int id = Convert.ToInt32(Request.QueryString["id"]);
        if(fromVal != "") 
        {
           from = Convert.ToDateTime(Request.QueryString["from"]);
        }
        if(toVal != "")
        {
           to = Convert.ToDateTime(Request.QueryString["to"]);
        }
        string sql = "SELECT * FROM DishonoredCheques WHERE InstanceTypeId = 11";
       


        try
        {

            if (id != 0 && fromVal != "" && toVal != "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE InstanceTypeId = '" + id + "' AND IssueDate >= '" + Convert.ToDateTime(Request.QueryString["from"]) + "' AND IssueDate <='" + Convert.ToDateTime(Request.QueryString["to"]) + "';";
            }
            else if (id == 0 && fromVal != "" && toVal != "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE IssueDate >= '" + Convert.ToDateTime(Request.QueryString["from"]) + "' AND IssueDate <='" + Convert.ToDateTime(Request.QueryString["to"]) + "';";

                // return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from && x.IssueDate <= to).ToList());
            }
            else if (id != 0 && fromVal == "" && toVal == "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE InstanceTypeId = '" + id + "'";

                // return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id == InstanceId).ToList());
            }
            else if (id == 0 && fromVal == "" && toVal == "")
            {
                sql = "";

                // return View(_context.DishonoredCheques.ToList());
            }
            else if (id == 0 && fromVal == "" && toVal != "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE IssueDate <='" + Convert.ToDateTime(Request.QueryString["to"]) + "';";

            }
            else if (id == 0 && fromVal != "" && toVal == "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE IssueDate >= '" + Convert.ToDateTime(Request.QueryString["from"]) + "';";

                // return View(_context.DishonoredCheques.Where(x => x.IssueDate >= from).ToList());
            }
            else if (id != 0 && fromVal == "" && toVal != "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE InstanceTypeId = '" + id + "' AND IssueDate <='" + Convert.ToDateTime(Request.QueryString["to"]) + "';";

                // return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate <= to).ToList());
            }
            else if (id != 0 && fromVal != "" && toVal == "")
            {
                sql = "SELECT * FROM DishonoredCheques WHERE InstanceTypeId = '" + id + "' AND IssueDate >= '" + Convert.ToDateTime(Request.QueryString["from"]) + "';";

                // return View(_context.DishonoredCheques.Where(x => x.InstanceType.Id >= InstanceId && x.IssueDate >= from).ToList());
            }

        }
        catch (Exception ex)
        {

        }
        finally
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["sqlConn"].ConnectionString))
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand(sql, connection);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                connection.Close();
                DataSet ds = new DataSet();
                sda.Fill(ds);
                ReportDocument rpt = new ReportDocument();
                rpt.Load(Server.MapPath("CrystalReport1.rpt"));
                rpt.SetDatabaseLogon("sa", "sola@9220", "HQITPROGLAP165", "DCMS");
                rpt.SetDataSource(ds.Tables["table"]);
                CrystalReportViewer1.ReportSource = rpt;
                CrystalReportViewer1.DisplayGroupTree = false;
                CrystalReportViewer1.DataBind();
            }
        }
    }
}
