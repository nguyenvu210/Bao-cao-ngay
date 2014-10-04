using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Reflection;

/// <summary>
/// Summary description for excel
/// </summary>
public class csv 
{
    public OleDbConnection con = null;
    public string str = "";
    public string filepath = "";
    //*** DataTable ***//

    public csv()
	{
		//
		// TODO: Add constructor logic here
		//
	}
    public void Open(string Delimited)
    {
        try
        {
            if (con == null)
            {
                str="Provider=Microsoft.Jet.OLEDB.2.0;Data Source=" + filepath + ";Extended Properties='text;HDR=Yes;FMT=Delimited("+Delimited+")';";
                con = new OleDbConnection(str);
                con.Open();
            }
        }
        catch (Exception )
        {
            str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties='text;HDR=Yes;FMT=Delimited(" + Delimited + ")';";
            con = new OleDbConnection(str);
            con.Open();
        }
    }
    public void Close()
    {
        if (con != null)
        {
            con.Close();
            con = null;
        }
    }

    public DataTable CsvCreateDataTable(String filename, string Delimited)
    {
        this.Open(Delimited);
        OleDbDataAdapter dtAdapter;
        DataTable dt = new DataTable();

        String strSQL;
        strSQL = "SELECT * FROM " + filename + "";

        dtAdapter = new OleDbDataAdapter(strSQL, con);
        dtAdapter.Fill(dt);

        dtAdapter = null;

        this.Close();

        return dt; //*** Return DataTable ***//
    }
    public System.Collections.ArrayList CsvCreateFile(string filename)
    {
        System.Collections.ArrayList arr = new System.Collections.ArrayList();
        StreamReader sr = new StreamReader(filename);
        string inputLine = "";
        while ((inputLine = sr.ReadLine()) != null)
        {
            arr.Add(inputLine);
        }
        sr.Close();
        return arr;
    }
    public DataTable TransferCSVToTable(string filePath, string Delimited)
    {
        DataTable dt = new DataTable();
        string[] csvRows = System.IO.File.ReadAllLines(filePath);
        string[] fields = null;
        for (int i = 0; i < csvRows.Length;i++)
        {
            fields = csvRows[i].Split(new string[] { Delimited }, StringSplitOptions.None);
            if (i == 0)
            {
                int count = 0;
                //DataColumn[] cols = new DataColumn[fields.Length];
                for (int j = 0; j < fields.Length; j++)
                {
                    try
                    {
                        //cols[j] = new DataColumn(fields[j].ToString().Trim(), typeof(string));
                        dt.Columns.Add(new DataColumn(fields[j].ToString().Trim(), typeof(string)));
                    }
                    catch (Exception)
                    {
                        count++;
                        //dt.Columns.AddRange(cols);
                        dt.Columns.Add(new DataColumn("EXISTED_" + count.ToString(), typeof(string)));
                    }
                }
            }
            else
            {                
                DataRow row = dt.NewRow();
                row.ItemArray = fields;
                dt.Rows.Add(row);
            }
        }
        return dt;
    }
}
