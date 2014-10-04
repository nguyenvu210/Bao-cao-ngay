using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
    class Funtion
    {
        public static string Convert_xls_to_xlsx(string filetemplate, string tenfile, string filein)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook Ewbook;
            Ewbook = xlApp.Workbooks.Open(filetemplate, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Ewbook.SaveAs(filein + "\\" + tenfile + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault,
                                null, null, false, false, Excel.XlSaveAsAccessMode.xlExclusive,
                                false, false, false, false, false);

            if (System.IO.File.Exists(filetemplate))
                System.IO.File.Delete(filetemplate);
            Ewbook = null;
            xlApp.Quit();
            return "Finish";
        }
        public static DataTable Combine(string sheet,string path,int header,int start)
        {
            DataTable temp = OpenExcelFilesheet.GetWorksheetSingle(sheet, path);
            DataTable Data = new DataTable();
           
            for (int ii = 0; ii < temp.Columns.Count; ii++)
            {
                if (header != 0)
                {
                    Data.Columns.Add(temp.Rows[header][ii].ToString());
                }
                else
                {
                    Data.Columns.Add(temp.Columns[ii].ColumnName.ToString());
                }
               
            }
            DataRow dtRow = Data.NewRow();
            for (int k = start; k < temp.Rows.Count; k++)
            {
                for (int h = 0; h < temp.Columns.Count; h++)
                {
                    string ds = temp.Rows[k][h].ToString();
                    if (ds == "")
                    {
                        dtRow[h] = 0;
                    }
                    else
                    {
                        dtRow[h] = temp.Rows[k][h];
                    }
                    if (temp.Rows[k][h].ToString().ToLower().IndexOf("label") != -1)
                    {
                        string a = temp.Rows[k][h].ToString();
                        a = a.Replace(",", " ");
                        dtRow[h] = a;
                    }
                }
                Data.Rows.Add(dtRow);
                dtRow = Data.NewRow();
            }
            return Data;
        }
        public static DataTable GetSheet(string Sheet, string path)
        {
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            OleDbConnection con = new System.Data.OleDb.OleDbConnection(connectionString);
            OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter(
                "select * from [" + Sheet + "$]", con);
            try
            {
                con.Open();
            }
            catch
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                con = new System.Data.OleDb.OleDbConnection(connectionString);
                cmd = new System.Data.OleDb.OleDbDataAdapter(
               "select * from [" + Sheet + "$]", con);
                con.Open();
            }
            DataSet excelDataSet = new DataSet();
            cmd.Fill(excelDataSet);
            con.Close();
            return excelDataSet.Tables[0];
        }
        public static void Export_to_CSV(DataTable data, string strFilePath, string duoi)
        {
            string dis = ",";
            string[] savetam = strFilePath.Split('.');
            strFilePath = savetam[0] + duoi;
            StreamWriter sw = new StreamWriter(strFilePath, false);
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

