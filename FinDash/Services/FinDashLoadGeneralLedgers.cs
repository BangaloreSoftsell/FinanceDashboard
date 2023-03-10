using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Web;
using FinDash.Constants;
using FinDash.Logger;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinDash.Services
{
    public class FinDashLoadGeneralLedgers
    {
        public void LoadGeneralLedger(string transactionNumber, int curMonth, int curYear, string strTime, DateTime timeReceived)
        {            
            SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
            con.Open();
            string sql = "select FilesAllowed.FileName,Files.Month,Files.Status,files.SchoolID,files.CreatedBy,files.CreatedOn,files.FileID,Schools.FilesPath from  FinDash.llac.Files, FinDash.llac.FilesAllowed ,FinDash.llac.Schools where Files.FileID = FilesAllowed.FilesAllowedID and Schools.SchoolID = FilesAllowed.SchoolID  and Files.Year='" + curYear + "' and Files.Month='" + curMonth + "' and Files.Status ='2' and  FilesAllowed.FileType = '1768'";
            SqlCommand sqlCommand = new SqlCommand(sql, con);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            if (sqlDataReader.HasRows)
            {
                while (sqlDataReader.Read())
                {
                    //LoadForm
                    string fileName = sqlDataReader.GetString(0);
                    var lastMonth = sqlDataReader.GetInt32(1);
                    var fileStatus = sqlDataReader.GetInt32(2);
                    var schoolId = sqlDataReader.GetInt32(3);
                    var createdBy = sqlDataReader.GetInt32(4);
                    DateTime createdOn = sqlDataReader.GetDateTime(5);
                    var fileId = sqlDataReader.GetInt32(6);
                    string targetDir = sqlDataReader.GetString(7);
                    if (fileStatus == 2)
                    {
                        //sqlDataReader.Close();
                        SqlCommand myUpdate = new SqlCommand("Update FinDash.llac.Files set Files.Status ='3' , ProcessingStarted='" + strTime + "' where Files.FileName='" + fileName + "' and Files.Status='" + fileStatus + "' and Files.Month='" + lastMonth + "'", con);
                        myUpdate.ExecuteNonQuery();
                        FinDashLogger finDashLogger = new FinDashLogger();
                        finDashLogger.Logger(transactionNumber, schoolId, "LoadGeneralLedgers()", "LoadGeneralLedgers()", "Updatating  the status is successfull ", timeReceived, createdBy, createdOn);                       
                    }
                    //Read Excel ColumnCount and Row Count
                    string[] fileEntries = Directory.GetFiles(targetDir, "*gl*");
                    foreach (string fileSName in fileEntries)
                    {
                        Excel.Application x1App = new Excel.Application();
                        Excel.Workbook x1wkb = x1App.Workbooks.Open(fileSName);
                        Excel.Worksheet x1wks = x1wkb.Sheets[1];
                        Excel.Range x1range = x1wks.UsedRange;
                        int rowCount = x1range.Rows.Count;
                        int colcountExcel = x1range.Columns.Count;
                        string sqlCount = "select SchoolID,FilesAllowed.ColumnsCount from FinDash.llac.FilesAllowed where FileType='1768' and SchoolID ='" + schoolId + "'";
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
                                    FinDashLogger finDashLogger = new FinDashLogger();
                                    finDashLogger.Logger(transactionNumber, schoolId, "LoadGeneralLedgers()", "LoadGeneralLedgers()", "Column Check is successfull ", timeReceived, createdBy, createdOn);                                   
                                    if (schID == 1)
                                    {
                                        //Stanza.GLTransactions
                                    }
                                    else if (schID == 2)
                                    {
                                        //L4LCharleston.GLTransactions
                                    }
                                    else if (schID == 3)
                                    {
                                        //For Michigan actual row starts from 6
                                        int rowBegin = 6;
                                        //var gLDataObject = new GLDataObject();
                                        SetEachRowForMichigan(rowBegin);                                       
                                    }
                                    else if (schID == 4)
                                    {
                                        //FHSColumbus.GLTransactions
                                    }
                                    else if (schID == 5)
                                    {
                                        //FHSCleveland.GLTransactions 
                                    }
                                    else if (schID == 6)
                                    {
                                        //USLC.GLTransactions
                                    }
                                    else if (schID == 7)
                                    {
                                        //L4LEdgewood.GLTransactions
                                    }
                                    else
                                    {
                                        //Invalid SchoolID
                                    }
                                }
                                else
                                {                                    
                                    FinDashLogger finDashLogger = new FinDashLogger();
                                    finDashLogger.Logger(transactionNumber, schoolId, "LoadGeneralLedgers()", "LoadGeneralLedgers()", "Column count does not match ", timeReceived, createdBy, createdOn);

                                    SqlCommand transError = new SqlCommand(
                                                       "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(2) + "','TransactionBegin()','" + sqlDataReader.GetInt32(5) + "','Check Failed:Column Count Does not Match -- Failed','" + timeReceived + "')", con);
                                    transError.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }
                sqlDataReader.Close();
            }
            else
            {               
                FinDashLogger finDashLogger = new FinDashLogger();
                finDashLogger.Logger(transactionNumber, sqlDataReader.GetInt32(3), "LoadGeneralLedgers()", "LoadGeneralLedgers()", "Updatating  the status is successfull ", timeReceived, sqlDataReader.GetInt32(4), sqlDataReader.GetDateTime(5));

                SqlCommand transError = new SqlCommand(
                                   "INSERT INTO [llac].[Error] (TransactionNumber, SchoolID,ErrorAt,FilesID,ErrorMessage,ErrorLoggedAt)" + "Values ('" + transactionNumber + "','" + sqlDataReader.GetInt32(2) + "','TransactionBegin()','" + sqlDataReader.GetInt32(5) + "','Insert Failed','" + timeReceived + "')", con);
                transError.ExecuteNonQuery();
            }
            con.Close();
        }
        private void SetEachRowForMichigan(int rowBegin)
        {
            try
            {
                SqlConnection con = new SqlConnection(Controller.Connections.DBConn);
                con.Open();
                string filePath = FinDashConstants.filesPath;
                string excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", filePath);
                Excel.Application x1App = new Excel.Application();
                Excel.Workbook x1wkb = x1App.Workbooks.Open(filePath);
                Excel.Worksheet x1wks = x1wkb.Sheets[1];
                Excel.Range x1range = x1wks.UsedRange;
                int rowCount = x1range.Rows.Count;
                int columnCount = x1range.Columns.Count;
                int i = 0;
                int j = 0;
                //GLDataObject glDataobj = new GLDataObject();
                object[,] glDataObject = x1range.Value2;
                for (i = rowBegin; i <= rowCount; i++)
                {
                    if (x1range.Cells[i, 1].Value2 == FinDashConstants.headerMichiganCell)
                    {
                        i++;
                    }
                    if (x1range.Cells[i, 1].Value2 == FinDashConstants.cellColBalance  || x1range.Cells[i, 1].Value2 == FinDashConstants.cellColGeneral)
                    {
                        glDataObject = (object[,])x1range.Rows[i, Type.Missing].Value;
                        string dateParse = glDataObject[1, 7].ToString();
                        DateTime dateOnly = DateTime.ParseExact(dateParse, "MM/dd/yy", CultureInfo.InvariantCulture);
                        glDataObject[1, 7] = dateOnly.ToString("dd/MM/yyyy");
                    }
                    string[] curRow = new string[columnCount];
                    for (j = 1; j <= columnCount; j++)
                    {
                        if (x1range.Cells[i, j].Value2 != null)
                        {
                            curRow[j] = x1range.Cells[i, j].Value2.ToString();
                        }
                        if (curRow[1] == FinDashConstants.cellColBalance || curRow[1] == FinDashConstants.cellColGeneral)
                        {
                            //setHeaderValueforRow(i, glDataObject);
                        }
                        else if (curRow[1] == null && curRow[4] != null && curRow[10] != null)
                        {
                            if (curRow[24] != null)
                            {
                                goto Account;
                            }
                            else if (curRow[25] != null)
                            {
                                goto Account;
                            }
                        }
                        else if (curRow[24] != null)
                        {
                            goto DebitCredit;
                        }
                        else if (curRow[25] != null)
                        {
                            goto DebitCredit;
                        }
                    }
                    if (curRow[7] != null)
                    {
                        DateTime dateOnly = DateTime.ParseExact(curRow[7].ToString(), "MM/dd/yy", CultureInfo.InvariantCulture);
                        curRow[7] = dateOnly.ToString("dd/MM/yyyy");
                    }
                    goto ConDB;
                Account:
                    //setAccountValueforRow(i, curRow);
                    goto ConDB;
                DebitCredit:
                    //SetDebitCreditDescription(i, curRow);
                    goto ConDB;
                ConDB:
                    if (curRow[1] == FinDashConstants.cellColBalance || curRow[1] == FinDashConstants.cellColGeneral)
                    {
                        SqlCommand myCommand = new SqlCommand(
                        "INSERT INTO [FHSMichigan].[GLTransactions] (JournalType, TransactionNumber, Date,Posted,TransactionDescription)" + "Values ('" + glDataObject[1, 1] + "','" + glDataObject[1, 6] + "','" + glDataObject[1, 7].ToString() + "','" + glDataObject[1, 8] + "','" + glDataObject[1, 9] + "')", con);
                        myCommand.ExecuteNonQuery();
                    }
                    else if (curRow[18] != null)
                    {
                        SqlCommand myCommand = new SqlCommand(
                        "INSERT INTO [FHSMichigan].[GLTransactions] (JournalType, AccountCode, TransactionNumber, Date,Posted,TransactionDescription, AccountDescription, Description, Debits)" + "Values ('" + glDataObject[1, 1] + "','" + curRow[4] + "','" + glDataObject[1, 6] + "','" + glDataObject[1, 7].ToString() + "','" + glDataObject[1, 8] + "','" + glDataObject[1, 9] + "','" + curRow[10] + "', '" + curRow[14] + "', '" + curRow[18] + "')", con);
                        myCommand.ExecuteNonQuery();
                    }
                    else if (curRow[24] != null)
                    {
                        if (curRow[11] != null)
                        {
                            SqlCommand myCommand = new SqlCommand(
                            "INSERT INTO [FHSMichigan].[GLTransactions] (JournalType, AccountCode, TransactionNumber, Date,Posted,TransactionDescription, AccountDescription, Description, Credits)" + "Values ('" + glDataObject[1, 1] + "','" + curRow[4] + "','" + glDataObject[1, 6] + "','" + glDataObject[1, 7].ToString() + "','" + glDataObject[1, 8] + "','" + glDataObject[1, 9] + "','" + curRow[10] + "', '" + curRow[11] + "', '" + curRow[24] + "')", con);
                            myCommand.ExecuteNonQuery();
                        }
                        else if (curRow[15] != null)
                        {
                            SqlCommand myCommand = new SqlCommand(
                            "INSERT INTO [FHSMichigan].[GLTransactions] (JournalType, AccountCode, TransactionNumber, Date,Posted,TransactionDescription, AccountDescription, Description, Credits)" + "Values ('" + glDataObject[1, 1] + "','" + curRow[4] + "','" + glDataObject[1, 6] + "','" + glDataObject[1, 7].ToString() + "','" + glDataObject[1, 8] + "','" + glDataObject[1, 9] + "','" + curRow[10] + "', '" + curRow[15] + "', '" + curRow[24] + "')", con);
                            myCommand.ExecuteNonQuery();
                        }
                    }
                    else if (curRow[25] != null)
                    {
                        SqlCommand myCommand = new SqlCommand(
                        "INSERT INTO [FHSMichigan].[GLTransactions] (JournalType, AccountCode, TransactionNumber, Date,Posted,TransactionDescription, AccountDescription, Description, Credits)" + "Values ('" + glDataObject[1, 1] + "','" + curRow[4] + "','" + glDataObject[1, 6] + "','" + glDataObject[1, 7].ToString() + "','" + glDataObject[1, 8] + "','" + glDataObject[1, 9] + "','" + curRow[10] + "', '" + curRow[15] + "', '" + curRow[25] + "')", con);
                        myCommand.ExecuteNonQuery();
                    }
                }
                con.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}