using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using FinDash.Services;

namespace FinDash.Controller
{
    public static class Connections
    {
        public static string DBConn = ConfigurationManager.ConnectionStrings["FinDashConnection"].ConnectionString;
    }
    public class FinDashController
    {
        FindashDataLoader finDashDataLoader = new FindashDataLoader();
        bool status { get; set; }

        public void FinDashDataLoader()
        {            
            status = finDashDataLoader.FilesCheck();                      
        }      
    }   
}