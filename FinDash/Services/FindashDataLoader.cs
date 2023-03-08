using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;
using FinDash.Constants;

namespace FinDash.Services
{
    public class FindashDataLoader
    {
        bool status { get; set; }
        string strTime = DateTime.Parse(DateTime.Now.ToString()).TimeOfDay.ToString();
        TimeSpan timeReceived = TimeSpan.Parse(DateTime.Parse(DateTime.Now.ToString()).TimeOfDay.ToString());
        public bool FilesCheck() 
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();                       
            var now = DateTime.Now.ToString("ddMMyyyy");
            int min = timeReceived.Minutes;
            int sec = timeReceived.Seconds;
            var transactionNumber = now + min + sec;
            //filePath instead of FTP folder
            //string filePath = @"C:\Users\sweet\Sweety\FHS Michigan\File";
            string filePath = FinDashConstants.filesPath;
            string[] fileEntries = Directory.GetFiles(filePath, "*balance*");
            int curMonth=0, curYear=0;
            foreach (string listFName in fileEntries)
            {
                string FileName = listFName.Substring(listFName.LastIndexOf("\\") + 1);
                curMonth = int.Parse(FileName.Substring(0, 2));
                curYear = int.Parse(FileName.Substring(6, 2));
                if (curYear >= 4)
                {
                    string stringYear = String.Format("{0:2000}", curYear);
                    curYear = int.Parse(stringYear);
                }
            }
            string sql = "Select FilesAllowed.FileName, Schools.FilesPath, FilesAllowed.SchoolID, FilesAllowed.CreatedBy, FilesAllowed.CreatedOn,FilesAllowed.FilesAllowedID from FinDash.llac.FilesAllowed ,FinDash.llac.Schools where Type='6' and Status='8'  and FilesAllowed.SchoolID=schools.SchoolID";
            SqlCommand sqlCommand = new SqlCommand(sql, con);
            sqlCommand.CommandTimeout = 0;
            using (SqlDataReader reader = sqlCommand.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string fileName = reader.GetString(0);
                        var schoolId = reader.GetInt32(2);
                        var createdBy = reader.GetInt32(3);
                        DateTime createdOn = reader.GetDateTime(4);
                        var fileID = reader.GetInt32(5);                        
                        // File name, Date to be added from FTP folder 
                        SqlCommand myCommand = new SqlCommand(
                                    "INSERT INTO [llac].[Files] (SchoolID, Year, Month,FileID,FileName,TimeReceived,Status,CreatedBy,CreatedOn)" + "Values ('" + schoolId + "','" + curYear + "','" + curMonth + "','" + fileID + "','" + fileName + "','" + timeReceived + "','2','" + createdBy + "','" + createdOn + "')", con);
                        myCommand.ExecuteNonQuery();
                        SqlCommand myLogger = new SqlCommand(
                                    "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schoolId + "','TransactionBegin()','TransactionBegin()','Insert statement successfull','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                        myLogger.ExecuteNonQuery();

                        ColsCheck(transactionNumber, curMonth, curYear);
                        FinDashLoadGeneralLedgers finDashLoadGeneralLedgers=new FinDashLoadGeneralLedgers();
                        finDashLoadGeneralLedgers.LoadGeneralLedger(transactionNumber, curMonth, curYear,strTime,timeReceived);                        
                    }
                }
                else
                {
                    SqlCommand myLogger = new SqlCommand(
                                   "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + reader.GetInt32(2) + "','TransactionBegin()','TransactionBegin()','Insert Failed','" + TimeSpan.Parse(strTime) + "','" + reader.GetInt32(3) + "','" + reader.GetTimeSpan(4) + "')", con);
                    myLogger.ExecuteNonQuery();
                    SqlCommand transError = new SqlCommand(
                                   "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + reader.GetInt32(2) + "','TransactionBegin()','" + reader.GetInt32(5) + "','Insert Failed','" + timeReceived + "')", con);
                    transError.ExecuteNonQuery();
                }
            }
            Console.WriteLine("Status of Files Check");
            //status=RowCheck();
            con.Close();
            return status;            
        }
        private void ColsCheck(string transactionNumber,int curMonth,int curYear)
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            string sql = "SELECT FilesAllowed.FileName,Files.Month,Files.Status,files.SchoolID,files.CreatedBy,files.CreatedOn,files.FileID,FilesAllowed.FileType,Schools.FilesPath,Files.Year from FinDash.llac.FilesAllowed ,FinDash.llac.Schools,FinDash.llac.Files where Files.FileID = FilesAllowed.FilesAllowedID and Files.SchoolID = Schools.SchoolID and Files.Year='" + curYear + "'and Files.Month = '" + curMonth + "' and Files.Status = '2' and  FilesAllowed.FileType in ('1765','1766','1770')";
            SqlCommand sqlCommand = new SqlCommand(sql, con);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            if (sqlDataReader.HasRows)
            {
                while (sqlDataReader.Read())
                {
                    string fileName = sqlDataReader.GetString(0);
                    var lastMonth = sqlDataReader.GetInt32(1);
                    var fileStatus = sqlDataReader.GetInt32(2);
                    var schoolId = sqlDataReader.GetInt32(3);
                    var createdBy = sqlDataReader.GetInt32(4);
                    DateTime createdOn = sqlDataReader.GetDateTime(5);
                    var fileId = sqlDataReader.GetInt32(6);
                    var fileType = sqlDataReader.GetInt32(7);
                    string targetDir = sqlDataReader.GetString(8);
                    var lastYear = sqlDataReader.GetInt32(9);
                    SqlCommand myUpdate = new SqlCommand("Update FinDash.llac.Files set Files.Status ='3' , ProcessingStarted='" + strTime + "' where Files.FileName='" + fileName + "' and Files.Status='" + fileStatus + "' and Files.Month='" + lastMonth + "'", con);
                    myUpdate.ExecuteNonQuery();
                    SqlCommand myLogger = new SqlCommand(
                                   "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schoolId + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Updatating the status is successfull','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                    myLogger.ExecuteNonQuery();
                    int sheetNum = 1;
                    if (fileType == 1770)
                    {
                        if (schoolId == 1)
                        {
                            //Open the second worksheet
                            sheetNum = 2;
                        }
                        else if (schoolId == 2)
                        {
                            //Open Sheet 5
                            sheetNum = 5;
                        }
                        else
                        {
                            //Open Sheet 1
                            sheetNum = 1;
                        }
                    }
                    if (fileName.Contains("alance") == true)
                    {
                        string[] fileEntries = Directory.GetFiles(targetDir, "*balance*");                        
                        foreach (string fileSName in fileEntries)
                        {
                            string FileName = fileSName.Substring(fileSName.LastIndexOf("\\") + 1);
                            SqlCommand myUpdatefile = new SqlCommand("Update FinDash.llac.Files set Files.FileName ='" + FileName + "' , ProcessingStarted='" + strTime + "' where Files.Status='3' and Files.Month='" + lastMonth + "'", con);
                            myUpdatefile.ExecuteNonQuery();
                            Excel.Application x1App = new Excel.Application();
                            Excel.Workbook x1wkb = x1App.Workbooks.Open(fileSName);
                            Excel.Worksheet x1wks = x1wkb.Sheets[sheetNum];
                            Excel.Range x1range = x1wks.UsedRange;
                            int rowCount = x1range.Rows.Count;
                            int colcountExcel = x1range.Columns.Count;
                            //the used range is calculated for the merged cell 'O'
                            colcountExcel = colcountExcel - 1;
                            string sqlCount = "select SchoolID,FilesAllowed.ColumnsCount from FinDash.llac.FilesAllowed where FileType='" + fileType + "' and SchoolID ='" + schoolId + "'";
                            SqlCommand sqlColcount = new SqlCommand(sqlCount, con);
                            SqlDataReader sqlDataCount = sqlColcount.ExecuteReader();
                            if (sqlDataCount.HasRows)
                            {
                                while (sqlDataCount.Read())
                                {
                                    var schID = sqlDataCount.GetInt32(0);
                                    var colCount = sqlDataCount.GetInt32(1);
                                    if (colcountExcel == colCount)
                                    {
                                        SqlCommand myLoggerExcel = new SqlCommand(
                                      "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schoolId + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Column Check is successfull','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                        myLoggerExcel.ExecuteNonQuery();
                                        SqlCommand myUpdateSuccess = new SqlCommand("Update FinDash.llac.Files set Files.Status ='5' , ProcessingStarted='" + strTime + "' where Files.FileName='" + fileName + "' and Files.Status='" + fileStatus + "' and Files.Month='" + lastMonth + "'", con);
                                        myUpdateSuccess.ExecuteNonQuery();
                                        RowsCheck(fileSName, schID, lastMonth, lastYear, sheetNum, transactionNumber);                                        
                                    }
                                    else
                                    {
                                        SqlCommand myLoggerExcel = new SqlCommand(
                                       "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(3) + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Check Failed-Column Count Does not Match','" + TimeSpan.Parse(strTime) + "','" + sqlDataReader.GetInt32(4) + "','" + sqlDataReader.GetTimeSpan(5) + "')", con);
                                        myLoggerExcel.ExecuteNonQuery();
                                        SqlCommand transError = new SqlCommand(
                                                           "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(2) + "','TransactionBegin()','" + sqlDataReader.GetInt32(5) + "','Check Failed:Column Count Does not Match -- Failed','" + timeReceived + "')", con);
                                        transError.ExecuteNonQuery();
                                        //Send Email regarding
                                    }
                                }
                            }
                        }
                    }
                    else if (fileName.Contains("PnL") == true)
                    {
                        //PnL
                        string[] fileEntries = Directory.GetFiles(targetDir, "*PnL*");
                        foreach (string fileSName in fileEntries)
                        {
                            Excel.Application x1App = new Excel.Application();
                            Excel.Workbook x1wkb = x1App.Workbooks.Open(fileSName);
                            Excel.Worksheet x1wks = x1wkb.Sheets[sheetNum];
                            Excel.Range x1range = x1wks.UsedRange;
                            int rowCount = x1range.Rows.Count;
                            int colcountExcel = x1range.Columns.Count;
                            string sqlCount = "select SchoolID,FilesAllowed.ColumnsCount from FinDash.llac.FilesAllowed where FileType='" + fileType + "' and SchoolID ='" + schoolId + "'";
                            SqlCommand sqlColcount = new SqlCommand(sqlCount, con);
                            SqlDataReader sqlDataCount = sqlColcount.ExecuteReader();
                            if (sqlDataCount.HasRows)
                            {
                                while (sqlDataCount.Read())
                                {
                                    var schID = sqlDataCount.GetInt32(0);
                                    var colCount = sqlDataCount.GetInt32(1);
                                    if (colcountExcel == colCount)
                                    {
                                        SqlCommand myLoggerExcel = new SqlCommand(
                                      "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schoolId + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Column Check is successfull','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                        myLoggerExcel.ExecuteNonQuery();
                                        SqlCommand myUpdateSuccess = new SqlCommand("Update FinDash.llac.Files set Files.Status ='5' , ProcessingStarted='" + strTime + "' where Files.FileName='" + fileName + "' and Files.Status='" + fileStatus + "' and Files.Month='" + lastMonth + "'", con);
                                        myUpdateSuccess.ExecuteNonQuery();
                                        RowsCheck(fileSName, schID, lastMonth, lastYear, sheetNum, transactionNumber);                                        
                                    }
                                    else
                                    {
                                        SqlCommand myLoggerExcel = new SqlCommand(
                                       "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(3) + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Check Failed-Column Count Does not Match','" + TimeSpan.Parse(strTime) + "','" + sqlDataReader.GetInt32(4) + "','" + sqlDataReader.GetTimeSpan(5) + "')", con);
                                        myLoggerExcel.ExecuteNonQuery();
                                        SqlCommand transError = new SqlCommand(
                                                           "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(2) + "','TransactionBegin()','" + sqlDataReader.GetInt32(5) + "','Check Failed:Column Count Does not Match -- Failed','" + timeReceived + "')", con);
                                        transError.ExecuteNonQuery();

                                        //mark that this should be part of the email sent to the Agency.
                                    }
                                }
                            }
                        }
                    }
                }
                sqlDataReader.Close();
            }
            else
            {
                SqlCommand myLogger = new SqlCommand(
                                  "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(3) + "','LoadGeneralLedgers()','LoadGeneralLedgers()','Update Failed-No Files in Load Ledgers','" + TimeSpan.Parse(strTime) + "','" + sqlDataReader.GetInt32(4) + "','" + sqlDataReader.GetTimeSpan(5) + "')", con);
                myLogger.ExecuteNonQuery();
                SqlCommand transError = new SqlCommand(
                                   "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(2) + "','TransactionBegin()','" + sqlDataReader.GetInt32(5) + "','Insert Failed','" + timeReceived + "')", con);
                transError.ExecuteNonQuery();
            }
            Console.WriteLine("Status of Row Check");                     
            con.Close();
        }
        private void RowsCheck(string Excelfile, int scId, int mon, int year, int sheetNum, string transactionNumber)
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            string excelPath = Path.GetDirectoryName(Excelfile);
            string sqlList = "select ColumnNumber,ColumnType from [FinDash].[llac].[ColumnsFormat] , [FinDash].[llac].[Files],[FinDash].[llac].Schools where Schools.SchoolID=ColumnsFormat.SchoolID and Schools.FilesPath='" + excelPath + "' and Files.Status='2' and files.FileID=FilesID";
            SqlCommand cmd = new SqlCommand(sqlList, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cont = 0;
            int[] ColumnNumber = new int[dt.Rows.Count];
            int[] ColumnType = new int[dt.Rows.Count];
            foreach (DataRow row in dt.Rows)
            {
                ColumnNumber[cont] = row.Field<int>(0);
                ColumnType[cont] = row.Field<int>(1);
                cont++;
            }
            Excel.Application x1App = new Excel.Application();
            Excel.Workbook x1wkb = x1App.Workbooks.Open(Excelfile);
            Excel.Worksheet x1wks = x1wkb.Sheets[sheetNum];
            Excel.Range x1range = x1wks.UsedRange;
            int rowCount = x1range.Rows.Count;
            int colcountExcel = x1range.Columns.Count;
            string sql = "select FilesID,Schools.SchoolID,RowCheckFrom ,ColumnNumber ,ColumnType,Schools.FilesPath,RangeFrom,RangeTo,ColumnsFormat.CreatedBy,ColumnsFormat.CreatedOn from [FinDash].[llac].[ColumnsFormat] , [FinDash].[llac].[Files],[FinDash].[llac].Schools where Schools.SchoolID=ColumnsFormat.SchoolID and Schools.FilesPath='" + excelPath + "' and Files.Status='2' and files.FileID=FilesID";
            SqlCommand sqlCommand = new SqlCommand(sql, con);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            while (sqlDataReader.Read())
            {
                var fileID = sqlDataReader.GetInt32(0);
                var schID = sqlDataReader.GetInt32(1);
                var rowCheck = sqlDataReader.GetInt32(2);
                var colType = sqlDataReader.GetInt32(4);
                var colNum = sqlDataReader.GetInt32(3);
                var rangeFrom = sqlDataReader.GetInt32(6);
                var rangeTo = sqlDataReader.GetInt32(7);
                var createdBy = sqlDataReader.GetInt32(8);
                DateTime createdOn = sqlDataReader.GetDateTime(9);
                string[] curRow = new string[colcountExcel];
                string strTime = DateTime.Parse(DateTime.Now.ToString()).TimeOfDay.ToString();
                var now = DateTime.Now.ToString("ddMMyyyy");
                var date = DateTime.Parse(DateTime.Now.ToString()).TimeOfDay.ToString();
                var seconds = TimeSpan.FromTicks(DateTime.Now.Ticks).TotalSeconds;
                TimeSpan timeReceived = TimeSpan.Parse(strTime);
                int Error = 0;
                int amt = 0;
                int m = 0;
                object[,] glDataObject = x1range.Value2;
                for (int i = rowCheck; i <= rowCount; i++)
                {
                    glDataObject = (object[,])x1range.Rows[i, Type.Missing].Value;
                    for (int j = 1; j < colcountExcel; j++)
                    {
                        if (x1range.Cells[i, j].Value2 == null)
                        {
                            curRow[j] = "Null";
                        }
                        else if (x1range.Cells[i, j].Value2 != null)
                        {
                            curRow[j] = x1range.Cells[i, j].Value2.ToString();

                            if (j == 1)
                            {
                                Regex regex = new Regex("[0-9]+");
                                Match match = regex.Match(curRow[j]);
                                if (!match.Success)
                                {
                                    Error++;
                                    goto Error;
                                }
                                else
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                            "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RowsCheck()','Regex Match is success','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                }
                            }
                            else if (j == 2 || j == 3 || j == 7)
                            {
                                Regex regex = new Regex("[A-Za-z]+");
                                Match match = regex.Match(curRow[j]);
                                if (!match.Success)
                                {
                                    Error++;
                                    goto Error;
                                }
                                else
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                            "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RowsCheck()','Regex Match is success','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                }
                            }
                            else if (j == 4 || j == 5 || j == 6 || j == 8 || j == 9 || j == 10 || j == 12 || j == 13 || j == 14)
                            {
                                Regex regex = new Regex("^(?![\\s\\S])+");
                                Match match = regex.Match(curRow[j]);
                                if (!match.Success)
                                {
                                    Error++;
                                    goto Error;
                                }
                                else
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                            "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RowsCheck()','Regex Match is success','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                    if (fileID == 1)
                                    {
                                        amt = x1range.Cells[i, 11].Value2;
                                    }
                                }
                            }
                            else if (j == 11)
                            {
                                Regex regex = new Regex(@"^-?[0-9][0-9,\.]+$");
                                Match match = regex.Match(curRow[j]);
                                if (!match.Success)
                                {
                                    Error++;
                                    goto Error;
                                }
                                else
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                            "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RowsCheck()','Regex Match is success','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                }

                                double curAmount = Convert.ToDouble(curRow[11]);
                                if (curAmount >= rangeFrom && curAmount <= rangeTo)
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                              "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RangeCheck()','Amount is within the range','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                }
                                else
                                {
                                    SqlCommand myLoggerExcel = new SqlCommand(
                              "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RangeCheck()','Amount is not within the range','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                    myLoggerExcel.ExecuteNonQuery();
                                    SqlCommand amtError = new SqlCommand(
                               "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + schID + "','178','" + fileID + "','Invalid Amount','" + timeReceived + "')", con);
                                    amtError.ExecuteNonQuery();
                                }
                            }
                        //TypeAndRangecheck(j, curRow, colType, ColumnNumber);
                        Error:
                            if (Error >= 1)
                            {
                                SqlCommand myLogger = new SqlCommand(
                                     "INSERT INTO [llac].[Logger] (TransactionNumber, SchoolID,ClassCalledFrom,FunctionCalledFrom,Message,LoggedOn,CreatedBy,CreatedOn)" + "Values ('" + transactionNumber + "','" + schID + "','RowsCheck()','RangeCheck()','Amount is not within the range','" + timeReceived + "','" + createdBy + "','" + createdOn + "')", con);
                                myLogger.ExecuteNonQuery();
                                SqlCommand transError = new SqlCommand(
                           "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + schID + "','181','" + fileID + "','Invalid Amount','" + timeReceived + "')", con);
                                transError.ExecuteNonQuery();
                                Error = 0;
                                //Email :Mark that this should be part of the email sent to the Agency.  
                            }
                        }
                    }
                    if (schID == 3)
                    {
                        int colsType = ColumnType[m];
                        if (fileID == 1)
                        {
                            if (Excelfile.Contains("alance") == true)
                            {
                                BalanceSheet(year, mon, colsType, createdBy, createdOn, curRow);
                            }
                        }
                        else if (fileID == 2)
                        {
                            if (Excelfile.Contains("PnL") == true)
                            {
                                RevenueAndExpenses(year, mon, colType, createdBy, createdOn, curRow);
                            }
                        }
                    }
                }
                //  goto DoneBalance;
            }
            con.Close();
            Console.WriteLine("Status of Range Check");           
        }
        public void BalanceSheet(int year, int mon, int colsType, int createdBy, DateTime createdOn, string[] glObject)
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            string query = "select ListId,AccountCode,AccountCodeDesc,ListItem from [FinDash].[llac].[Lists] where Lists.SchoolID='3'";
            SqlCommand cmd = new SqlCommand(query, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int cont = 0;
            int[] ListId = new int[dt.Rows.Count];
            string[] AccountCode = new string[dt.Rows.Count];
            string[] AccountCodeDesc = new string[dt.Rows.Count];
            string[] ListItem = new string[dt.Rows.Count];
            foreach (DataRow row in dt.Rows)
            {
                ListId[cont] = row.Field<int>(0);
                AccountCode[cont] = row.Field<string>(1);
                AccountCodeDesc[cont] = row.Field<string>(2);
                ListItem[cont] = row.Field<string>(2);
                cont++;
            }
            if (glObject[1] != "Null")
            {
                var check = Array.IndexOf(AccountCode, glObject[1]);
                if (check > -1)
                {
                    colsType = ListId[check];
                }
                if (glObject[11] == "Null")
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();

                }
                else
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();
                }
            }
            else if (glObject[1] == "Null" && glObject[3] != "Null")
            {
                var check = Array.IndexOf(AccountCodeDesc, glObject[3]);
                if (check > -1)
                {
                    colsType = ListId[check];
                }
                if (glObject[11] == "Null")
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();

                }
                else
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();
                }
            }
            else if (glObject[1] == "Null" && glObject[2] != "Null")
            {
                var check = Array.IndexOf(AccountCodeDesc, glObject[2]);
                if (check > -1)
                {
                    colsType = ListId[check];
                }
                if (glObject[11] == "Null")
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();

                }
                else
                {
                    SqlCommand cmdAssests = new SqlCommand(
                                       "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                    cmdAssests.ExecuteNonQuery();
                }
            }
            else
            {
                //Error and logger Unmatched List for school
                SqlCommand cmdAssests = new SqlCommand(
                                      "INSERT INTO [FinDash].[FHSMichigan].[AssetsAndLiabilities] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','1','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                cmdAssests.ExecuteNonQuery();
            }
            con.Close();
        }
        public void RevenueAndExpenses(int year, int mon, int colsType, int createdBy, DateTime createdOn, string[] glObject)
        {
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            if (glObject[1] != "Null")
            {
                string query = "select ListId,AccountCode,AccountCodeDesc,ListItem from [FinDash].[llac].[Lists] where Lists.SchoolID='3'";
                SqlCommand cmd = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                int cont = 0;
                int[] ListId = new int[dt.Rows.Count];
                string[] AccountCode = new string[dt.Rows.Count];
                string[] AccountCodeDesc = new string[dt.Rows.Count];
                string[] ListItem = new string[dt.Rows.Count];
                foreach (DataRow row in dt.Rows)
                {
                    ListId[cont] = row.Field<int>(0);
                    AccountCode[cont] = row.Field<string>(1);
                    AccountCodeDesc[cont] = row.Field<string>(2);
                    ListItem[cont] = row.Field<string>(3);
                    cont++;
                }
                if (glObject[1] != "Null")
                {
                    var check = Array.IndexOf(AccountCode, glObject[1]);
                    if (check > -1)
                    {
                        colsType = ListId[check];
                        if (ListItem[check] == "Revenue")
                        {
                            if (glObject[9] == "Null" && glObject[10] == "Null" && glObject[11] == "Null" && glObject[17] == "Null" && glObject[20] == "Null")
                            {
                                SqlCommand cmdRevenue = new SqlCommand(
                                                 "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdRevenue.ExecuteNonQuery();
                            }
                            else if (glObject[9] != "Null" && glObject[10] != "Null" && glObject[17] != "Null" && glObject[20] != "Null")
                            {
                                SqlCommand cmdRevenue = new SqlCommand(
                                                "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,Budget,Actual,Variance,BudgetPercentage,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[10] + "','" + glObject[17] + "','" + glObject[20] + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdRevenue.ExecuteNonQuery();
                            }
                            else if (glObject[11] != "Null" && glObject[10] == "Null" && glObject[9] != "Null" && glObject[17] != "Null" && glObject[19] != "Null")
                            {
                                SqlCommand cmdRevenue = new SqlCommand(
                                                "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,Budget,Actual,Variance,BudgetPercentage,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[11] + "','" + glObject[17] + "','" + glObject[20] + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdRevenue.ExecuteNonQuery();
                            }
                            else if (glObject[10] == "Null" && glObject[11] == "Null" && glObject[9] != "Null")
                            {
                                if (glObject[20] == "Null")
                                {
                                    SqlCommand cmdRevenue = new SqlCommand(
                                                 "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,Budget,Variance,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[17] + "','" + createdBy + "','" + createdOn + "')", con);
                                    cmdRevenue.ExecuteNonQuery();
                                }
                            }
                            else
                            {

                            }
                        }
                        else if (ListItem[check] == "Expenditure")
                        {
                            if (glObject[9] == "Null" && glObject[10] == "Null" && glObject[11] == "Null" && glObject[17] == "Null" && glObject[20] == "Null")
                            {
                                SqlCommand cmdExpenses = new SqlCommand(
                                    "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdExpenses.ExecuteNonQuery();
                            }
                            else if (glObject[10] != "Null" && glObject[9] != "Null" && glObject[17] != "Null" && glObject[20] != "Null")
                            {
                                SqlCommand cmdRevenue = new SqlCommand(
                                                "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,Budget,Actual,Variance,BudgetPercentage,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[10] + "','" + glObject[17] + "','" + glObject[20] + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdRevenue.ExecuteNonQuery();
                            }
                            else if (glObject[11] != "Null" && glObject[10] == "Null" && glObject[9] != "Null" && glObject[17] != "Null" && glObject[20] != "Null")
                            {
                                SqlCommand cmdRevenue = new SqlCommand(
                                                "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,Budget,Actual,Variance,BudgetPercentage,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[11] + "','" + glObject[17] + "','" + glObject[20] + "','" + createdBy + "','" + createdOn + "')", con);
                                cmdRevenue.ExecuteNonQuery();
                            }
                            else if (glObject[10] == "Null" && glObject[11] == "Null" && glObject[9] != "Null")
                            {
                                if (glObject[20] == "Null")
                                {
                                    SqlCommand cmdRevenue = new SqlCommand(
                                                 "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,Budget,Variance,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[9] + "','" + glObject[17] + "','" + createdBy + "','" + createdOn + "')", con);
                                    cmdRevenue.ExecuteNonQuery();
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                }
                else if (glObject[1] == "Null" && glObject[3] != "Null")
                {
                    var check = Array.IndexOf(AccountCodeDesc, glObject[3]);
                    if (check > -1)
                    {
                        colsType = ListId[check];
                    }
                    if (glObject[11] == "Null")
                    {
                        SqlCommand cmdRevenue = new SqlCommand(
                                           "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                        cmdRevenue.ExecuteNonQuery();
                        SqlCommand cmdExpenses = new SqlCommand(
                                          "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + createdBy + "','" + createdOn + "')", con);
                        cmdRevenue.ExecuteNonQuery();

                    }
                    else
                    {
                        SqlCommand cmdAssests = new SqlCommand(
                                           "INSERT INTO [FinDash].[FHSMichigan].[Revenue] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                        cmdAssests.ExecuteNonQuery();
                        SqlCommand cmdExpenses = new SqlCommand(
                                         "INSERT INTO [FinDash].[FHSMichigan].[Expenses] (SchoolID,Year,Month,Type,Value,CreatedBy,CreatedOn)" + "Values ('3','" + year + "','" + mon + "','" + colsType + "','" + glObject[11] + "','" + createdBy + "','" + createdOn + "')", con);
                        cmdExpenses.ExecuteNonQuery();
                    }
                }
                else
                {
                    //Error and logger Unmatched List for school
                }
            }
            con.Close();
        }        
    }
}