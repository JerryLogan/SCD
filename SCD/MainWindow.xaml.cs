using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Data;
using System.Net;
namespace SCD
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();    
        }
        string version = "v1.0.0";
        string filename;
        int speed_num = 0; //總共
        int index = 0;
        string[,] scdinfo= new string[2000,15];
        string SCD_DATABASE_PATH = "";
        private DatabaseStructure SCD_DATABASE;
        string DB_Item = " (CityName, RegionName, Address, DeptNm, BranchNm, Longitude, Latitude, direct, direct2, lmt, NSInd, EWInd, Altitue, Type, Note) ";
        public enum TSDB_ITEM
        {
            CityName = 0,
            RegionName,
            Address,
            DeptNm,
            BranchNm,
            Longitude = 5,
            Latitude,
            direct,
            direct2,
            lmt,
            NSInd = 10,
            EWInd,
            Altitude,
            Type,
            Note = 14,
            TS_MAX
        }
        string[] MPData_Result = new string[ (Int32)TSDB_ITEM.TS_MAX];
        private void Button_Click(object sender, RoutedEventArgs e)//load btn
        {

            Console.WriteLine("open file");
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            
            //dlg.FileName = "SCD"; // Default file name
            dlg.DefaultExt = ".xls |  .xlsx"; // Default file extension
            dlg.Filter = "*.xls |  *.xlsx"; // Filter files by extension
            
            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                filename = dlg.FileName;
            }
            Console.WriteLine("file="+ filename);
            lable1.Visibility = Visibility.Visible;
            lable1.Content = filename;
            var excelFile = new LinqToExcel.ExcelQueryFactory(filename);
            var sheet1 = excelFile.Worksheet("deleteType3");


            foreach (var itemRow in sheet1)
            {
                speed_num+=1;
                /*
                Console.WriteLine(speed_num+"：{0}、：{1}、：{2}、 : {3} 、 : {4} 、: {5}",
                    itemRow["CityName"],
                    itemRow["RegionName"],
                    itemRow["DeptNm"],
                    itemRow["Longitude"],
                    itemRow["Latitude"],
                    itemRow["direct"]
                    );
                */
                if (itemRow["CityName"] == "設置縣市")
                    continue;
#if false
                scdinfo[index, 0] = Convert.ToDouble(itemRow["Longitude"]);
                scdinfo[index, 1] = Convert.ToDouble(itemRow["Latitude"]);
                scdinfo[index, 3] = Convert.ToDouble(itemRow["direct2"]);
                scdinfo[index, 4] = Convert.ToDouble(itemRow["limit"]);
                scdinfo[index, 5] = Convert.ToDouble(itemRow["NSInd"]);
                scdinfo[index, 6] = Convert.ToDouble(itemRow["EWInd"]);
                scdinfo[index, 7] = Convert.ToDouble(itemRow["Altitue"]);
                scdinfo[index, 8] = Convert.ToDouble(itemRow["Type"]);
#else
//                string DB_Item = " (CityName, RegionName, Address, DeptNm, BranchNm, Longitude, Latitude, direct, direct2, lmt, NSInd, EWInd, Altitue, Type, Note) ";
                scdinfo[index, (Int32)TSDB_ITEM.CityName] = (itemRow["CityName"]);
                scdinfo[index, (Int32)TSDB_ITEM.RegionName] = (itemRow["RegionName"]);
                scdinfo[index, (Int32)TSDB_ITEM.Address] = (itemRow["Address"]);
                scdinfo[index, (Int32)TSDB_ITEM.DeptNm] = (itemRow["DeptNm"]);
                scdinfo[index, (Int32)TSDB_ITEM.BranchNm] = (itemRow["BranchNm"]);
                scdinfo[index, (Int32)TSDB_ITEM.Longitude] = (itemRow["Longitude"]);
                scdinfo[index, (Int32)TSDB_ITEM.Latitude] = (itemRow["Latitude"]);
                scdinfo[index, (Int32)TSDB_ITEM.direct] = (itemRow["direct"]);
                scdinfo[index, (Int32)TSDB_ITEM.direct2] = (itemRow["direct2"]);
                scdinfo[index, (Int32)TSDB_ITEM.lmt] = (itemRow["lmt"]);
                scdinfo[index, (Int32)TSDB_ITEM.NSInd] = (itemRow["NSInd"]);
                scdinfo[index, (Int32)TSDB_ITEM.EWInd] = (itemRow["EWInd"]);
                scdinfo[index, (Int32)TSDB_ITEM.Altitude] = (itemRow["Altitude"]);
                scdinfo[index, (Int32)TSDB_ITEM.Type] = (itemRow["Type"]);
                scdinfo[index, (Int32)TSDB_ITEM.Note] = (itemRow["Note"]);
#endif
                index++;
            }
            speed_num -= 1;
            Console.WriteLine("總共" + speed_num + "支");
            label2.Visibility = Visibility.Visible;
            label2.Content = "總共" + speed_num + "支";

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)//convert btn
        {
            SCD_DATABASE_PATH = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\SCD_Data.db";
            SCD_DATABASE = new DatabaseStructure(SCD_DATABASE_PATH);
            DBSCDCreateTable();
#if true
            #region 寫入SCDinfo檔
            label2.Content = speed_num + "支 轉換中...";
            int i = 0;
            for (i = 0; i < speed_num; i++)
            {
                MPData_Result[(Int32)TSDB_ITEM.CityName] = scdinfo[i, (Int32)TSDB_ITEM.CityName].ToString();
                MPData_Result[(Int32)TSDB_ITEM.RegionName] = scdinfo[i, (Int32)TSDB_ITEM.RegionName].ToString();
                MPData_Result[(Int32)TSDB_ITEM.Address] = scdinfo[i, (Int32)TSDB_ITEM.Address].ToString();
                MPData_Result[(Int32)TSDB_ITEM.DeptNm] = scdinfo[i, (Int32)TSDB_ITEM.DeptNm].ToString();
                MPData_Result[(Int32)TSDB_ITEM.BranchNm] = scdinfo[i, (Int32)TSDB_ITEM.BranchNm].ToString();
                MPData_Result[(Int32)TSDB_ITEM.direct] = scdinfo[i, (Int32)TSDB_ITEM.direct].ToString();
                MPData_Result[(Int32)TSDB_ITEM.Altitude] = scdinfo[i, (Int32)TSDB_ITEM.Altitude].ToString();
                MPData_Result[(Int32)TSDB_ITEM.Note] = scdinfo[i, (Int32)TSDB_ITEM.Note].ToString();

                PrintLog("SCDInfo.SCD[" + i + "].SC_LONGITUDE =" + scdinfo[i, (Int32)TSDB_ITEM.Longitude] +";");//經
                MPData_Result[(Int32)TSDB_ITEM.Longitude] = scdinfo[i, (Int32)TSDB_ITEM.Longitude].ToString();

                PrintLog("SCDInfo.SCD[" + i + "].SC_EWInd =" + "2" + ";");//EW
                MPData_Result[(Int32)TSDB_ITEM.EWInd] = "2";

                PrintLog("SCDInfo.SCD[" + i + "].SC_LATITUDE =" + scdinfo[i, (Int32)TSDB_ITEM.Latitude] +";");//緯
                MPData_Result[(Int32)TSDB_ITEM.Latitude] = scdinfo[i, (Int32)TSDB_ITEM.Latitude].ToString();

                PrintLog("SCDInfo.SCD[" + i + "].SC_NSInd =" + "1" + ";");//NS
                MPData_Result[(Int32)TSDB_ITEM.NSInd] = "1";

                PrintLog("SCDInfo.SCD[" + i + "].DIRECTION = SC_DIRECTION_" + scdinfo[i, (Int32)TSDB_ITEM.direct2] + ";");//方
                MPData_Result[(Int32)TSDB_ITEM.direct2] = scdinfo[i, (Int32)TSDB_ITEM.direct2].ToString();

                PrintLog("SCDInfo.SCD[" + i + "].LIMIT = SPEED_CAM_LIMIT_0" + scdinfo[i, (Int32)TSDB_ITEM.lmt] + ";");//限
                MPData_Result[(Int32)TSDB_ITEM.lmt] = scdinfo[i, (Int32)TSDB_ITEM.lmt].ToString();

                PrintLog("SCDInfo.SCD[" + i + "].TYPE = SC_Type_" + scdinfo[i, (Int32)TSDB_ITEM.Type] + ";");//種
                MPData_Result[(Int32)TSDB_ITEM.Type] = scdinfo[i, (Int32)TSDB_ITEM.Type].ToString();

                PrintLog("");
                Console.WriteLine("i=" + i  );
                DBInsertData();
            }
            label2.Content =   speed_num + "支 完成";
#endregion
#endif
            
        }

        public static void PrintLog(string pMessage)
        {
            string tFileDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);//@"C:\Users\jerry_lo\Desktop";
            string tFilePath = tFileDir + "\\SCDinfo.txt";

            StreamWriter tStreamWriter = null;
            try
            {
                if (!Directory.Exists(tFileDir))
                {
                    Directory.CreateDirectory(tFileDir);
                }
                if (!File.Exists(tFilePath))
                {
                    File.Create(tFilePath).Close();
                }
                tStreamWriter = new StreamWriter(tFilePath, true, System.Text.UTF8Encoding.UTF8);
                pMessage =  pMessage + "\r\n";
                tStreamWriter.Write(pMessage);
                tStreamWriter.Flush();
                tStreamWriter.Close();
            }
            catch (Exception)
            {
                //LogManager.PrintLog("PrintLog Exception 2");
            }

            GC.Collect();
        }

        public bool DBSCDCreateTable()
        {
            DateTime Date = DateTime.Now;
            string cmdStr;
            //string preserve = "jj123";

            cmdStr = "CREATE TABLE IF NOT EXISTS " + "SCD"  + DB_Item + ";";
            SCD_DATABASE.SQLiteCmdExecution(SCD_DATABASE_PATH, cmdStr);

            return true;
        }
        public void DBInsertData()
        {
            DateTime Date = DateTime.Now;

            string valuestr = "'" + MPData_Result[0] + "'";
            string insertstr;
            for (int index = 1; index < (Int32)TSDB_ITEM.TS_MAX; index++)
                valuestr = valuestr + ", " + "'" + MPData_Result[index] + "'";

            insertstr = "INSERT INTO " + "SCD" + DB_Item + "VALUES (" + valuestr + ");";

            SCD_DATABASE.SQLiteCmdExecution(SCD_DATABASE_PATH, insertstr);            
        }
    }


#region Database Function
    class DatabaseStructure
    {
        public DatabaseStructure()
        {
            CreateSQLiteDatabase("MPData.db");
        }

        public DatabaseStructure(string Database)
        {
            CreateSQLiteDatabase(Database);
        }

        // Data base connection
        public SQLiteConnection OpenConn(string Database)
        {
            //LogManager.PrintLog("Database connection: " + Database);
            string cnstr = string.Format("Data Source= " + Database + ";Version=3;New=False;Compress=True;");
            SQLiteConnection icn = new SQLiteConnection();
            icn.ConnectionString = cnstr;

            int TSCount = 0;
            while ((TSCount < 50) && (icn.State == ConnectionState.Open))
            {
                TSCount++;
                System.Threading.Thread.Sleep(10);
            }
            //LogManager.PrintLog("Rollback icn.State: " + icn.State + " TSCount " + TSCount);
            if (icn.State == ConnectionState.Open)
                icn.Close();
            icn.Open();
            return icn;
        }
        // Create data base
        public void CreateSQLiteDatabase(string Database)
        {
            //LogManager.PrintLog("Create Database: " + Database);
            string cnstr = string.Format("Data Source= " + Database + ";Version=3;New=False;Compress=True;");
            SQLiteConnection icn = new SQLiteConnection();
            icn.ConnectionString = cnstr;
            icn.Open();
            icn.Close();
        }

        //Execute Sqlite command
        public void SQLiteCmdExecution(string Database, string SqlSelectString)
        {
            SQLiteConnection icn = OpenConn(Database);
            SQLiteCommand cmd = new SQLiteCommand(SqlSelectString, icn);
            SQLiteTransaction mySqlTransaction = icn.BeginTransaction();

            cmd.Transaction = mySqlTransaction;
            cmd.ExecuteNonQuery();
            mySqlTransaction.Commit();
            /*
            try
            {
                cmd.Transaction = mySqlTransaction;
                cmd.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (SqlException ex)  // 使用 SqlException
            {
                StringBuilder errorMessages = new StringBuilder();
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                //LogManager.PrintLog("SQL errorMessages: " + errorMessages);
                //LogManager.PrintLog("SqlSelectString: " + SqlSelectString);
                MessageBox.Show(errorMessages.ToString());
                mySqlTransaction.Rollback();
            }
            catch (Exception e)
            {
                //LogManager.PrintLog("Rollback icn.State: " + icn.State);
                //LogManager.PrintLog("SQLiteCmdExecution: " + e.ToString());
                //LogManager.PrintLog("SqlSelectString: " + SqlSelectString);
                MessageBox.Show("AAA");
                mySqlTransaction.Rollback();
            }*/
            if (icn.State == ConnectionState.Open)
                icn.Close();
        }
        //Execute Sqlite command
        public void SQLiteCmdExecution(string Database, string SqlSelectString, int dev)//for PRODUCEDATA debug
        {
            //LogManager.PrintLog("Execute Sqlite command: " + Database + " command: " + SqlSelectString);
            SQLiteConnection icn = OpenConn(Database);
            SQLiteCommand cmd = new SQLiteCommand(SqlSelectString, icn);
            SQLiteTransaction mySqlTransaction = icn.BeginTransaction();

            //bool isPRODUCE = string.Equals(Database, globalVarManager.ProduceDataBasePath + "\\ProduceData.db");
            bool isReplaceInto = SqlSelectString.Contains("REPLACE INTO ");
            //MessageBox.Show(isPRODUCE.ToString()+"\r"+ SqlSelectString+"\ris replace into:"+ isReplaceInto);

            try
            {
                cmd.Transaction = mySqlTransaction;
                cmd.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (SqlException ex)  // 使用 SqlException
            {
                StringBuilder errorMessages = new StringBuilder();
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                //LogManager.PrintLog("SQL errorMessages: " + errorMessages);
                //LogManager.PrintLog("SqlSelectString: " + SqlSelectString);
                //MessageBox.Show(errorMessages.ToString());
                mySqlTransaction.Rollback();
            }
            catch (Exception e)
            {
                //LogManager.PrintLog("Rollback icn.State: " + icn.State);
                //LogManager.PrintLog("SQLiteCmdExecution: " + e.ToString());
                //LogManager.PrintLog("SqlSelectString: " + SqlSelectString);
                mySqlTransaction.Rollback();
            }
            if (icn.State == ConnectionState.Open)
                icn.Close();
        }

        // Return Sqlite command result
        public bool SQLiteCmdResultExist(string Database, string SqlSelectString)
        {
            //globalFunction.PrintLog("Return Sqlite command result: " + Database + " command: " + SqlSelectString);
            int count = -1;
            SQLiteConnection icn = OpenConn(Database);
            SQLiteCommand cmd = new SQLiteCommand(SqlSelectString, icn);
            SQLiteTransaction mySqlTransaction = icn.BeginTransaction();
            try
            {
                cmd.Transaction = mySqlTransaction;
                count = Convert.ToInt32(cmd.ExecuteScalar());
                mySqlTransaction.Commit();
            }
            catch (Exception ex)
            {
                mySqlTransaction.Rollback();
                throw (ex);
            }
            if (icn.State == ConnectionState.Open)
                icn.Close();

            if (count >= 1)
                return true;
            else
                return false;  //This case is for no data
        }
    }
#endregion
   
}
