using FinDash.Controller;
using FinDash.Services;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FinDash
{
    public partial class _Default : Page
    {
      
        protected void Page_Load(object sender, EventArgs e)
        {
            Console.WriteLine(Controller.Connections.DBConn);
        }
        
        protected void DataLoader_Click(object sender, EventArgs e)
        {
            //OleDbConnection oleDbConnection=new OleDbConnection(Controller.Connections.DBConn);
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('msg');", true);
            FinDashController finDashController = new FinDashController();
            finDashController.FinDashDataLoader();
            con.Close();
        }

        protected void DataTransformer_Click(object sender, EventArgs e)
        {

        }

        protected void DataDisplay_Click(object sender, EventArgs e)
        {

        }
    }
}