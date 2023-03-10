using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace FinDash.Logger
{
    public class FinDashLogger
    {
        public void Logger(string transactionNumber,int schoolId,string classFrom,string functionFrom,string message,DateTime time,int createdBy,DateTime createdOn)
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            SqlCommand myLogger = new SqlCommand(
                      "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schoolId + "','"+ classFrom + "','"+ functionFrom +"','"+message +"','" + time + "','" + createdBy + "','" + createdOn + "')", con);
            myLogger.ExecuteNonQuery();          
            con.Close();
        }       
    }
}