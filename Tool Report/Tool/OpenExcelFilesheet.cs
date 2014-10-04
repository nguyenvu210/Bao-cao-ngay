using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;

    public class OpenExcelFilesheet
    {
        public static string[] OpenExcelFile(string Path, bool isOpenXMLFormat)
        {
            string[] workSheetNames = new string[] { };
            string connectionString;
            OleDbConnection con;
            if (isOpenXMLFormat)
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            else
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            con = new OleDbConnection(connectionString);
            con.Open();
            System.Data.DataTable dataSet = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            workSheetNames = new String[dataSet.Rows.Count];
            int i = 0;
            foreach (DataRow row in dataSet.Rows)
            {
                string a = row["TABLE_NAME"].ToString().Trim();
                a = a.Replace("'", "");
                int chieudaiduongdan = a.Trim().LastIndexOf("$");
                try
                {
                    workSheetNames[i] = a.Substring(0, chieudaiduongdan);
                }
                catch
                {

                }
                i++;
            }
            string aa = "";
            i = 0;
            int g = 0;
            string[] workSheet = new string[workSheetNames.Length];
            foreach (string t in workSheetNames)
            {
                if (t != null)
                {
                    if (aa != t.ToString())
                    {
                        aa = t.ToString();
                        workSheet[g] = t.ToString();
                        g++;
                    }
                    i++;
                }
            }
            if (con != null)
            {
                con.Close();
                con.Dispose();
            }
            if (dataSet != null)
                dataSet.Dispose();
            return workSheet;
        }
        public static DataSet GetWorksheet(string[] worksheetName, string path)
        {
            string connectionString = "";
            DataSet excelDataSet = new DataSet();
            string[] splitByDots = path.Split(new char[1] { '.' });
            if (splitByDots[1] == "xlsx")
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            }
            if (splitByDots[1] == "xls")
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            }
          
            if (connectionString != "")
            {
                OleDbConnection con = new System.Data.OleDb.OleDbConnection(connectionString);               
                con.Open();
                foreach (string sheet in worksheetName)
                {
                    //excelDataSet = new DataSet(sheet);
                    OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter(
                       "select * from [" + sheet + "$]", con);
                    cmd.Fill(excelDataSet,sheet);
                }
                con.Close();
                
            }
            else
            {
                MessageBox.Show("Can not open, only accpet .xlsx .xls");
            }
            return excelDataSet;
           
        }
        public static DataTable GetWorksheetSingle(string worksheetName, string path)
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES;\""; 
            OleDbConnection con = new System.Data.OleDb.OleDbConnection(connectionString);
            OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter(
                "select * from [" + worksheetName + "$]", con);
            try
            {
                con.Open();
            }
            catch
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                con = new System.Data.OleDb.OleDbConnection(connectionString);
                cmd = new System.Data.OleDb.OleDbDataAdapter(
               "select * from [" + worksheetName + "$]", con);
                con.Open();
            }
            DataSet excelDataSet = new DataSet();
            cmd.Fill(excelDataSet);
            con.Close();
            return excelDataSet.Tables[0];
        }
        public static DataTable GetText(string duongdan)
        {
           
            int tongchieudai = duongdan.Trim().Length;
            int chieudaiduongdan = duongdan.Trim().LastIndexOf("\\");
            string tenfile = duongdan.Trim().Substring(chieudaiduongdan, tongchieudai - chieudaiduongdan);
            tenfile = tenfile.Remove(0, 1).Trim();
            string tenduongdan = duongdan.Trim().Substring(0, chieudaiduongdan);
            DataTable data = new DataTable();
            string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + tenduongdan.Trim() + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
            string sql_select;
            System.Data.Odbc.OdbcConnection conn;
            conn = new System.Data.Odbc.OdbcConnection(strConnString.Trim());
            conn.Open();
            sql_select = "select * from [" + tenfile + "]";
            System.Data.Odbc.OdbcDataAdapter obj_oledb_da;
            obj_oledb_da = new System.Data.Odbc.OdbcDataAdapter(sql_select, conn);
            obj_oledb_da.Fill(data);
            conn.Close();
            return data;
        }
        public static void Export_to_CSV(DataTable data, string strFilePath,string duoi,string type)
        {
            string dis = ",";
            if (duoi == ".cel")
            {
                dis = "\t";
            }
            if (duoi == ".txt")
            {
                dis = "\t";
            }
            string[] savetam = strFilePath.Split('.');
            strFilePath = savetam[0]+ duoi;
            StreamWriter sw = new StreamWriter(strFilePath, false);
            if (type == "2")
            {
                sw.Write("2 TEMS_-_Cell_names");
                sw.Write(sw.NewLine);
            }
            if (type == "3")
            {
                sw.Write("2010 TEMS_-_Cell_names");
                sw.Write(sw.NewLine);
            }
            int iColCount = data.Columns.Count;
            for (int i = 0; i < iColCount; i++)
            {
                sw.Write(data.Columns[i]);
                if (i < iColCount - 1)
                {
                    sw.Write(dis);
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in data.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        sw.Write(dr[i].ToString());
                    }

                    if (i < iColCount - 1)
                    {
                        sw.Write(dis);
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();

           
        }
    }

