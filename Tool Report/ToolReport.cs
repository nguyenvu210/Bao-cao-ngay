using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing.Text;
namespace Tool_Report
{
    public partial class ToolReport : Form
    {
        public ToolReport()
        {
            InitializeComponent();
        }
        FolderBrowserDialog BrowserDialog = new FolderBrowserDialog();
        string filein = "";
        string fileout = "";
        string template = "";
        int row_index_nokia = 0;
        int row_index_zte = 0;
        int row_index_huawei = 0;
        int row_index_nokia_3G = 0;
        int row_index_huawei_3G = 0;
        string Title_1 = "---------Ðang đọc dữ liệu từ file : ";
        string Title_2 = "--------------Ðang xử lý : ";
        DataTable dt_ProMap = new DataTable();
        DataTable PSR_LUSR = new DataTable();

        DataTable Daily_Report_Normal_1 = new DataTable("Daily_Report_Normal_1");
        DataTable Daily_Report_Normal_2 = new DataTable("Daily_Report_Normal_2");
        DataTable Daily_Report_Normal_3 = new DataTable("Daily_Report_Normal_3");
        DataTable Daily_Report_VQI_1 = new DataTable("Daily_Report_VQI_1");
        DataTable Daily_Report_Peak_1 = new DataTable("Daily_Report_Peak_1");
        DataTable Daily_Report_Peak_2 = new DataTable("Daily_Report_Peak_2");
        DataTable Daily_Report_Rx_Quality_1 = new DataTable("Daily_Report_Rx_Quality_1");
        DataTable GPRS_DATA_HW = new DataTable("GPRS_DATA_HW");

        DataSet HW = new DataSet();

        DataTable KPI_Normal_2G = new DataTable();    
        DataTable KPI_Peak_2G = new DataTable("KPI_Peak_2G");
        DataTable KPI_Peak_3G = new DataTable("KPI_Peak_3G");
        DataTable KPI_Normal_3G = new DataTable("KPI_Normal_3G");

        DataTable BTS_WS_BSC_Global_Cellday = new DataTable("BTS_WS_BSC_Global_Cellday");
        DataTable BTS_WS_BSC_Global_Cellbh = new DataTable("BTS_WS_BSC_Global_Cellbh");       
        DataTable KPI_cell_sua_cellbh = new DataTable("KPI_cell_sua_cellbh");
        DataTable KPI_cell_sua_cellday = new DataTable("KPI_cell_sua_cellday");
        DataTable Traffic_cellbh = new DataTable("Traffic_cellbh");
        DataTable Traffic_cellday = new DataTable();
        DataTable ULDL_ZONEDAY = new DataTable("ULDL_ZONEDAY");
        DataTable ULDL_PLMNDAY = new DataTable("ULDL_PLMNDAY");
        DataTable DISTRIBUTE_ZONEDAY = new DataTable("DISTRIBUTE_ZONEDAY");
        DataTable DISTRIBUTE_PLMNDAY = new DataTable("DISTRIBUTE_PLMNDAY");
        DataTable GPRS_DATA_NSN_1 = new DataTable("Packet_control_unit_measurement_(GPRS)");
        DataTable GPRS_DATA_NSN = new DataTable("(E)GPRS_KPIs_(226)");
        DataSet NSN = new DataSet();

        DataTable Query_KPI_Cell_CS_Normal = new DataTable("Query_KPI_Cell_CS_Normal");
        DataTable Query_KPI_Cell_CS_Peak = new DataTable("Query_KPI_Cell_CS_Peak");
        DataTable Query_KPI_Cell_GPRS = new DataTable("Query_KPI_Cell_GPRS");
        DataSet ZTE = new DataSet();


        DataTable KPI_Cell_Normal = new DataTable();
        DataTable KPI_Cell_Peak_CS = new DataTable();
        DataTable KPI_Cell_Peak_PS = new DataTable();
        DataTable KPI_Cell_Peak = new DataTable();
        DataSet HW_3G = new DataSet();


        DataTable VTRAN_A2_new_Viettel_day = new DataTable();
        DataTable VTRAN_A2_new_Viettel_bh = new DataTable();
        DataSet NSN_3G = new DataSet();

        DataTable LU_PER_LAC_NETACT3G = new DataTable();
        DataTable Report_paging = new DataTable();
        DataTable LUSR_2g_3g = new DataTable();
        DataTable Paging_2g_3g = new DataTable();
        DataSet Lusr_page = new DataSet();

        DataTable Rowlist = new DataTable();
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook Ewbook;
        Excel.Worksheet xlSht;
        Excel.Range Rows; 
        public void add_log_file(string noidung) // Viet noi dung vao loffile
        {
            lst_logfile.Items.Add(noidung.Trim());
            lst_logfile.SelectedIndex = lst_logfile.Items.Count - 1;
            Application.DoEvents();
        }
        private void Data_Click(object sender, EventArgs e)
        {
            if (BrowserDialog.ShowDialog(this) == DialogResult.Cancel)
                return;
            add_log_file("Data in:" + BrowserDialog.SelectedPath);
            filein = BrowserDialog.SelectedPath;
        }
        private void Template_Click(object sender, EventArgs e)
        {
            if (BrowserDialog.ShowDialog(this) == DialogResult.Cancel)
                return;
            add_log_file("Template :" + BrowserDialog.SelectedPath);
            template = BrowserDialog.SelectedPath;
        }

        private void Run_Click(object sender, EventArgs e)
        {
            if (BrowserDialog.ShowDialog(this) == DialogResult.Cancel)
                return;
            add_log_file("Data out:" + BrowserDialog.SelectedPath);
            fileout = BrowserDialog.SelectedPath;
            //-------------------------------------------------------
            string Timestart = DateTime.Now.ToString();
            string[] thoigian = Regex.Split(DateTime.Now.ToString(), @"\W+");
            string date = thoigian[1] + "_" + thoigian[0] + "_" + thoigian[2];
            Convert();
            if (DR2G.Checked)
            {
                add_log_file("---------------------------------------------------------------- Daily Report 2G ------------ Movitel, S.A --------------------------------------------------------");
                Loaddata();               
                Merge_data();
                KPI_Normal_2G = CongthucKPI.KPI_Normal_2G(NSN, HW, ZTE);             
                KPI_Peak_2G = CongthucKPI.KPI_Peak_2G(NSN, HW, ZTE);              
                OpenExcelFilesheet.Export_to_CSV(KPI_Normal_2G, fileout + "\\ 2G_KPI_Normal " + date, ".csv", "0");
                OpenExcelFilesheet.Export_to_CSV(KPI_Peak_2G, fileout + "\\ 2G_KPI_Peak " + date, ".csv", "0");
                DR_2G();
            }
            if (DR3G.Checked)
            {
                add_log_file("---------------------------------------------------------------- Daily Report 3G ------------ Movitel, S.A -------------------------");
                Loaddata();
                Merge_data();
                Filt3G();
                OpenExcelFilesheet.Export_to_CSV(KPI_Normal_3G, fileout + "\\ 3G_KPI_Normal_" + date, ".csv", "0");
                OpenExcelFilesheet.Export_to_CSV(KPI_Peak_3G, fileout + "\\ 3G_KPI_Peak_" + date, ".csv", "0");               
                DR_3G();
            }
            MessageBox.Show("Start : at " + Timestart + "\n" + "Finish : at " + DateTime.Now.ToString(), "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            add_log_file("-----------------------------------------Copyright by Actiontwo, ManhNV3, ---------------------------------Code design by Cuongnh7@viettel.com.vn");

        }
        private void Convert()
        {
            DirectoryInfo d = new DirectoryInfo(filein);
            FileInfo[] f = d.GetFiles();
            f = d.GetFiles();
            for (int i = 0; i < f.Length; i++)
            {
                string tenfile = f[i].Name;
                string[] check = tenfile.Split(new char[1] { '.' });
                tenfile = tenfile.Substring(0, tenfile.Length - 4);
                if (check[check.Length - 1] == "xls" || check[check.Length - 1] == "csv")
                {
                    add_log_file("Convert to .xlsx: " + f[i].Name);
                    Funtion.Convert_xls_to_xlsx(filein + "\\" + f[i].Name, tenfile, filein);
                }
            }
        }
        private void Merge_data()
        {
            if (DR2G.Checked)
            {
                Daily_Report_Normal_1.Clone();
                Daily_Report_Normal_1.PrimaryKey = new[] { Daily_Report_Normal_1.Columns[3] };
                Daily_Report_Normal_2.PrimaryKey = new[] { Daily_Report_Normal_2.Columns[3] };
                Daily_Report_Normal_3.PrimaryKey = new[] { Daily_Report_Normal_3.Columns[3] };
                Daily_Report_Normal_1.Merge(Daily_Report_Normal_3); Daily_Report_Normal_1.Merge(Daily_Report_Normal_2);

                Daily_Report_Peak_1.Clone();
                Daily_Report_Peak_1.PrimaryKey = new[] { Daily_Report_Peak_1.Columns[4] };
                Daily_Report_Peak_2.PrimaryKey = new[] { Daily_Report_Peak_2.Columns[4] };
                Daily_Report_Peak_1.Merge(Daily_Report_Peak_2);

                HW.Tables.Add(Daily_Report_Normal_1); HW.Tables[0].TableName = "Normal";
                HW.Tables.Add(Daily_Report_Peak_1); HW.Tables[1].TableName = "Peak";
                HW.Tables.Add(Daily_Report_VQI_1); HW.Tables[2].TableName = "VQI";
                HW.Tables.Add(Daily_Report_Rx_Quality_1); HW.Tables[3].TableName = "Rx_Quality";
                HW.Tables.Add(GPRS_DATA_HW); HW.Tables[4].TableName = "GPRS";

                BTS_WS_BSC_Global_Cellbh.Clone();
                BTS_WS_BSC_Global_Cellbh.PrimaryKey = new[] { BTS_WS_BSC_Global_Cellbh.Columns[3] };
                Traffic_cellbh.PrimaryKey = new[] { Traffic_cellbh.Columns[3] };
                KPI_cell_sua_cellbh.PrimaryKey = new[] { KPI_cell_sua_cellbh.Columns[3] };
                BTS_WS_BSC_Global_Cellbh.Merge(KPI_cell_sua_cellbh);
                BTS_WS_BSC_Global_Cellbh.Merge(Traffic_cellbh);


                BTS_WS_BSC_Global_Cellday.Clone();
                BTS_WS_BSC_Global_Cellday.PrimaryKey = new[] { BTS_WS_BSC_Global_Cellday.Columns[3] };
                Traffic_cellday.PrimaryKey = new[] { Traffic_cellday.Columns[3] };
                KPI_cell_sua_cellday.PrimaryKey = new[] { KPI_cell_sua_cellday.Columns[3] };
                BTS_WS_BSC_Global_Cellday.Merge(KPI_cell_sua_cellday);
                BTS_WS_BSC_Global_Cellday.Merge(Traffic_cellday);


                NSN.Tables.Add(BTS_WS_BSC_Global_Cellday); NSN.Tables[0].TableName = "Normal";
                NSN.Tables.Add(BTS_WS_BSC_Global_Cellbh); NSN.Tables[1].TableName = "Peak";
                NSN.Tables.Add(ULDL_ZONEDAY); NSN.Tables[2].TableName = "ULDL_ZONEDAY";
                NSN.Tables.Add(ULDL_PLMNDAY); NSN.Tables[3].TableName = "ULDL_PLMNDAY";
                NSN.Tables.Add(DISTRIBUTE_ZONEDAY); NSN.Tables[4].TableName = "DISTRIBUTE_ZONEDAY";
                NSN.Tables.Add(DISTRIBUTE_PLMNDAY); NSN.Tables[5].TableName = "DISTRIBUTE_PLMNDAY";
                NSN.Tables.Add(GPRS_DATA_NSN); NSN.Tables[6].TableName = "GPRS";
                NSN.Tables.Add(GPRS_DATA_NSN_1); NSN.Tables[7].TableName = "GPRS_1";

                ZTE.Tables.Add(Query_KPI_Cell_CS_Normal); ZTE.Tables[0].TableName = "Normal";
                ZTE.Tables.Add(Query_KPI_Cell_CS_Peak); ZTE.Tables[1].TableName = "Peak";
                ZTE.Tables.Add(Query_KPI_Cell_GPRS); ZTE.Tables[2].TableName = "GPRS";
                Lusr_page = new DataSet();
                Lusr_page.Tables.Add(LU_PER_LAC_NETACT3G); Lusr_page.Tables[0].TableName = "LU_PER_LAC_NETACT3G";
                Lusr_page.Tables.Add(Report_paging); Lusr_page.Tables[1].TableName = "Report_paging";
                Lusr_page.Tables.Add(LUSR_2g_3g); Lusr_page.Tables[2].TableName = "LUSR_2g_3g";
                Lusr_page.Tables.Add(Paging_2g_3g); Lusr_page.Tables[3].TableName = "Paging_2g_3g";
                Lusr_page.Tables.Add(PSR_LUSR); Lusr_page.Tables[4].TableName = "PSR_LUSR";
            }
            if (DR3G.Checked)
            {
                HW_3G.Tables.Add(KPI_Cell_Normal); HW_3G.Tables[0].TableName = "Normal";
                HW_3G.Tables.Add(KPI_Cell_Peak_CS); HW_3G.Tables[1].TableName = "Peak_CS";
                HW_3G.Tables.Add(KPI_Cell_Peak_PS); HW_3G.Tables[2].TableName = "Peak_PS";
                HW_3G.Tables.Add(KPI_Cell_Peak); HW_3G.Tables[3].TableName = "Peak";
                NSN_3G.Tables.Add(VTRAN_A2_new_Viettel_day); NSN_3G.Tables[0].TableName = "Normal";
                NSN_3G.Tables.Add(VTRAN_A2_new_Viettel_bh); NSN_3G.Tables[1].TableName = "Peak";
                Lusr_page = new DataSet();
                Lusr_page.Tables.Add(Report_paging); Lusr_page.Tables[0].TableName = "Report_paging";             
                Lusr_page.Tables.Add(Paging_2g_3g); Lusr_page.Tables[1].TableName = "Paging_2g_3g";
                Lusr_page.Tables.Add(PSR_LUSR); Lusr_page.Tables[2].TableName = "PSR_LUSR";
                
            }
        }
        private void Loaddata()
        {
            add_log_file("ProTemplate.xlsx");
            PSR_LUSR = Funtion.GetSheet("PSR_LUSR", template + "\\" + "ProTemplate.xlsx");
            PSR_LUSR = Funtion.Combine("PSR_LUSR", template + "\\" + "ProTemplate.xlsx", 0, 0);
            Rowlist = Funtion.GetSheet("Row", template + "\\" + "ProTemplate.xlsx");
            row_index_nokia = int.Parse(Rowlist.Rows[1][1].ToString());
            row_index_zte = int.Parse(Rowlist.Rows[2][1].ToString());
            row_index_huawei = int.Parse(Rowlist.Rows[0][1].ToString());
            row_index_nokia_3G = int.Parse(Rowlist.Rows[4][1].ToString());
            row_index_huawei_3G = int.Parse(Rowlist.Rows[3][1].ToString());

            if (DR2G.Checked)
            {
                #region 2G               
                dt_ProMap = OpenExcelFilesheet.GetWorksheetSingle("ProList2G", template + "\\" + "ProTemplate.xlsx");
                DirectoryInfo d = new DirectoryInfo(filein);
                FileInfo[] f = d.GetFiles();
                DataTable temp= new DataTable();
                for (int i = 0; i < f.Length; i++)
                {
                    string name = f[i].Name.Trim().ToLower();
                    string path = filein + "\\" + f[i].Name;
                    if (name.IndexOf("daily report normal_1") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Normal_1 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report normal_2") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Normal_2 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report normal_3") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Normal_3 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report vqi_1") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_VQI_1 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report peak_1") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Peak_1 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report peak_2") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Peak_2 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("daily report rx quality_1") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Daily_Report_Rx_Quality_1 = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("gprs data") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        GPRS_DATA_HW = Funtion.Combine("Sheet1", path, 8, 9);
                    }

                    if (name.IndexOf("query-_kpi cell cs normal") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Query_KPI_Cell_CS_Normal = Funtion.Combine("sheet1", path, 4, 5);
                    }
                    if (name.IndexOf("query-_kpi cell cs peak") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Query_KPI_Cell_CS_Peak = Funtion.Combine("sheet1", path, 4, 5);
                    }
                    if (name.IndexOf("query-_kpi cell gprs") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Query_KPI_Cell_GPRS = Funtion.Combine("sheet1", path, 4, 5);
                    }
                    if (name.IndexOf("bts,_ws,_bsc_&_global cellbh") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        BTS_WS_BSC_Global_Cellbh = Funtion.Combine("Data", path,0,1);
                    }
                    if (name.IndexOf("bts,_ws,_bsc_&_global cellday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        BTS_WS_BSC_Global_Cellday = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("traffic.") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Traffic_cellday = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("ul_and_dl_quality_per_trx_(197) zoneday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        ULDL_ZONEDAY = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("ul_and_dl_quality_per_trx_(197) plmnday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        ULDL_PLMNDAY = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("distribution_of_call_samples_by_codecs_and_quality_classes_(fer)_(245) zoneday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        DISTRIBUTE_ZONEDAY = Funtion.Combine("Data", path, 1, 1);
                    }
                    if (name.IndexOf("distribution_of_call_samples_by_codecs_and_quality_classes_(fer)_(245) plmnday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        DISTRIBUTE_PLMNDAY = Funtion.Combine("Data", path, 1, 1);
                    }
                    if (name.IndexOf("kpi_cell_sua cellbh") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_cell_sua_cellbh = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("kpi_cell_sua cellday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_cell_sua_cellday = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("traffic cellbh") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Traffic_cellbh = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("(e)gprs_kpis_(226)") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        GPRS_DATA_NSN = Funtion.Combine("Data", path, 1, 2);
                    }
                    if (name.IndexOf("packet_control_unit_measurement_(gprs)") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        GPRS_DATA_NSN_1 = Funtion.Combine("Data", path, 1, 2);
                    }

                    if (name.IndexOf("lu_per_lac_netact3g") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        LU_PER_LAC_NETACT3G = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.Trim().ToLower().IndexOf("report_paging") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Report_paging = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("lusr") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        LUSR_2g_3g = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                    if (name.IndexOf("paging 2g,3g") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Paging_2g_3g = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                }
                #endregion
            }
            if (DR3G.Checked)
            {
                dt_ProMap = OpenExcelFilesheet.GetWorksheetSingle("ProList3G", template + "\\" + "ProTemplate.xlsx");
                DirectoryInfo d = new DirectoryInfo(filein);
                FileInfo[] f = d.GetFiles();
                DataTable temp= new DataTable();
                for (int i = 0; i < f.Length; i++)
                {
                    string name = f[i].Name.Trim().ToLower();
                    string path = filein + "\\" + f[i].Name;
                    if (name.IndexOf("kpi cell normal") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_Cell_Normal = Funtion.Combine("KPI Cell Normal", path, 0, 0);                       
                    }
                    if (name.IndexOf("kpi cell peak cs") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_Cell_Peak_CS = Funtion.Combine("KPI Cell Peak CS", path, 0, 0);
                    }
                    if (name.IndexOf("kpi cell peak ps") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_Cell_Peak_PS = Funtion.Combine("KPI Cell Normal", path, 0, 0);
                    }
                    if (name.IndexOf("kpi cell peak.") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        KPI_Cell_Peak = Funtion.Combine("KPI Cell Peak", path, 0, 0);
                    }
                    if (name.IndexOf("vtran_a2_new_-_viettel_daily_report_new_(for_report) cellbh") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        VTRAN_A2_new_Viettel_bh = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("vtran_a2_new_-_viettel_daily_report_new_(for_report) cellday") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        VTRAN_A2_new_Viettel_day = Funtion.Combine("Data", path, 0, 1);
                    }
                    if (name.IndexOf("report_paging") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Report_paging = Funtion.Combine("Data", path, 0, 1);
                    }                   
                    if (f[i].Name.Trim().ToLower().IndexOf("paging 2g,3g") != -1)
                    {
                        add_log_file(Title_1 + f[i].Name);
                        Paging_2g_3g = Funtion.Combine("Sheet1", path, 8, 9);
                    }
                } 
               
            }
        }       
        private void Filt3G()
        {
           // -----------------------------------------
            KPI_Normal_3G = normal_peak_3G(KPI_Cell_Normal, VTRAN_A2_new_Viettel_day, "normal");
            add_log_file(" Finish Normal_3G");
            KPI_Peak_3G = normal_peak_3G(KPI_Cell_Peak, VTRAN_A2_new_Viettel_bh, "peak");
            add_log_file(" Finish Peak_3G");
        }       
       
        private DataTable normal_peak_3G(DataTable Huawei, DataTable Nokia,string Gio)
        {
            float G = 0; float H = 0; float I = 0; float J = 0; float K = 0; float L = 0; float M = 0; float N = 0; float O = 0; float P = 0; float Q = 0; float R = 0; float S = 0;
            float V = 0; float W = 0; float X = 0; float Y = 0; float Z = 0; float AA = 0; float AB = 0; float AC = 0; float AD = 0; float AE = 0; float AF = 0; float AG = 0; float AH = 0; float AI = 0;
            float AJ = 0; float AK = 0; float AO = 0; float BA = 0; float AX = 0; float AM = 0; float AR = 0; float AQ = 0; float AU = 0; float AV = 0; float AT = 0; float AW = 0; float AZ = 0;
            string A = ""; string B = ""; string C = ""; string D = ""; string E = ""; string F = "";
            DataRow dtRow = Huawei.NewRow();
            if (Gio == "normal")
            {
                for (int i = 0; i < Huawei.Rows.Count; i++)
                {
                    if (Huawei.Rows[i][6].ToString() != "") { Huawei.Rows[i][6] = float.Parse(Huawei.Rows[i][6].ToString()) * 24; } else { Huawei.Rows[i][6] = 0; }
                    if (Huawei.Rows[i][7].ToString() != "") { Huawei.Rows[i][7] = float.Parse(Huawei.Rows[i][7].ToString()) * 24; } else { Huawei.Rows[i][7] = 0; }
                   
                }
            }

            for (int i = 0; i < Nokia.Rows.Count; i++)
            {
                
                string sitename = Nokia.Rows[i][3].ToString();
                if (sitename != "" && sitename.Length > 6)
                {
                    A = Nokia.Rows[i][0].ToString();
                    B = Nokia.Rows[i][1].ToString();
                    C = Nokia.Rows[i][2].ToString();
                    D = Nokia.Rows[i][3].ToString();
                    E = Nokia.Rows[i][4].ToString();
                    F = Nokia.Rows[i][5].ToString();
                    if (Nokia.Rows[i][7].ToString() != "") { H = float.Parse(Nokia.Rows[i][7].ToString()); } else { H = 0; }
                    if (Nokia.Rows[i][8].ToString() != "") { I = float.Parse(Nokia.Rows[i][8].ToString()); } else { I = 0; }
                    if (Nokia.Rows[i][9].ToString() != "") { J = float.Parse(Nokia.Rows[i][9].ToString()); } else { J = 0; }
                    if (Nokia.Rows[i][10].ToString() != "") { K = float.Parse(Nokia.Rows[i][10].ToString()); } else { K = 0; }
                    if (Nokia.Rows[i][11].ToString() != "") { L = float.Parse(Nokia.Rows[i][11].ToString()); } else { L = 0; }
                    if (Nokia.Rows[i][12].ToString() != "") { M = float.Parse(Nokia.Rows[i][12].ToString()); } else { M = 0; }
                    if (Nokia.Rows[i][13].ToString() != "") { N = float.Parse(Nokia.Rows[i][13].ToString()); } else { N = 0; }
                    if (Nokia.Rows[i][14].ToString() != "") { O = float.Parse(Nokia.Rows[i][14].ToString()); } else { O = 0; }
                    if (Nokia.Rows[i][15].ToString() != "") { P = float.Parse(Nokia.Rows[i][15].ToString()); } else { P = 0; }
                    if (Nokia.Rows[i][16].ToString() != "") { Q = float.Parse(Nokia.Rows[i][16].ToString()); } else { Q = 0; }
                    if (Nokia.Rows[i][17].ToString() != "") { R = float.Parse(Nokia.Rows[i][17].ToString()); } else { R = 0; }
                    if (Nokia.Rows[i][18].ToString() != "") { S = float.Parse(Nokia.Rows[i][18].ToString()); } else { S = 0; }
                    if (Nokia.Rows[i][21].ToString() != "") { V = float.Parse(Nokia.Rows[i][21].ToString()); } else { V = 0; }
                    if (Nokia.Rows[i][22].ToString() != "") { W = float.Parse(Nokia.Rows[i][22].ToString()); } else { W = 0; }
                    if (Nokia.Rows[i][23].ToString() != "") { X = float.Parse(Nokia.Rows[i][23].ToString()); } else { X = 0; }
                    if (Nokia.Rows[i][24].ToString() != "") { Y = float.Parse(Nokia.Rows[i][24].ToString()); } else { Y = 0; }
                    if (Nokia.Rows[i][25].ToString() != "") { Z = float.Parse(Nokia.Rows[i][25].ToString()); } else { Z = 0; }
                    if (Nokia.Rows[i][26].ToString() != "") { AA = float.Parse(Nokia.Rows[i][26].ToString()); } else { AA = 0; }
                    if (Nokia.Rows[i][27].ToString() != "") { AB = float.Parse(Nokia.Rows[i][27].ToString()); } else { AB = 0; }
                    if (Nokia.Rows[i][28].ToString() != "") { AC = float.Parse(Nokia.Rows[i][28].ToString()); } else { AC = 0; }
                    if (Nokia.Rows[i][29].ToString() != "") { AD = float.Parse(Nokia.Rows[i][29].ToString()); } else { AD = 0; }
                    if (Nokia.Rows[i][30].ToString() != "") { AE = float.Parse(Nokia.Rows[i][30].ToString()); } else { AE = 0; }
                    if (Nokia.Rows[i][31].ToString() != "") { AF = float.Parse(Nokia.Rows[i][31].ToString()); } else { AF = 0; }
                    if (Nokia.Rows[i][32].ToString() != "") { AG = float.Parse(Nokia.Rows[i][32].ToString()); } else { AG = 0; }
                    if (Nokia.Rows[i][33].ToString() != "") { AH = float.Parse(Nokia.Rows[i][33].ToString()); } else { AH = 0; }
                    if (Nokia.Rows[i][34].ToString() != "") { AI = float.Parse(Nokia.Rows[i][34].ToString()); } else { AI = 0; }
                    if (Nokia.Rows[i][35].ToString() != "") { AJ = float.Parse(Nokia.Rows[i][35].ToString()); } else { AJ = 0; }
                    if (Nokia.Rows[i][36].ToString() != "") { AK = float.Parse(Nokia.Rows[i][36].ToString()); } else { AK = 0; }
                    if (Nokia.Rows[i][38].ToString() != "") { AM = float.Parse(Nokia.Rows[i][38].ToString()); } else { AM = 0; }
                    if (Nokia.Rows[i][40].ToString() != "") { AO = float.Parse(Nokia.Rows[i][40].ToString()); } else { AO = 0; }
                    if (Nokia.Rows[i][42].ToString() != "") { AQ = float.Parse(Nokia.Rows[i][42].ToString()); } else { AQ = 0; }
                    if (Nokia.Rows[i][43].ToString() != "") { AR = float.Parse(Nokia.Rows[i][43].ToString()); } else { AR = 0; }
                    if (Nokia.Rows[i][45].ToString() != "") { AT = float.Parse(Nokia.Rows[i][45].ToString()); } else { AT = 0; }
                    if (Nokia.Rows[i][46].ToString() != "") { AU = float.Parse(Nokia.Rows[i][46].ToString()); } else { AU = 0; }
                    if (Nokia.Rows[i][47].ToString() != "") { AV = float.Parse(Nokia.Rows[i][47].ToString()); } else { AV = 0; }
                    if (Nokia.Rows[i][48].ToString() != "") { AW = float.Parse(Nokia.Rows[i][48].ToString()); } else { AW = 0; }
                    if (Nokia.Rows[i][49].ToString() != "") { AX = float.Parse(Nokia.Rows[i][49].ToString()); } else { AX = 0; }
                    if (Nokia.Rows[i][51].ToString() != "") { AZ = float.Parse(Nokia.Rows[i][51].ToString()); } else { AZ = 0; }
                    if (Nokia.Rows[i][52].ToString() != "") { BA = float.Parse(Nokia.Rows[i][52].ToString()); } else { BA = 0; }


                    dtRow[0] = Nokia.Rows[i][0].ToString();
                    dtRow[1] = C;
                    dtRow[2] = F;
                    dtRow[3] = E;
                    dtRow[4] = D;
                    dtRow[5] = G;
                    dtRow[6] = H;
                    dtRow[7] = I;
                    dtRow[8] = N;
                    dtRow[9] = BA;
                    dtRow[10] = AO;
                    dtRow[11] = AX;
                    dtRow[12] = AM;
                    dtRow[13] = AR;
                    dtRow[14] = AQ;
                    dtRow[15] = AU;
                    dtRow[16] = AT;
                    dtRow[17] = AW;
                    dtRow[18] = AZ;
                    dtRow[19] = AW / AX * AQ / AR * 100;
                    dtRow[20] = AZ / BA * AT / AU * 100;
                    dtRow[21] = V;
                    dtRow[22] = Math.Round(W * V / 100, 0);
                    dtRow[23] = X;
                    dtRow[24] = Math.Round(X * Y / 100, 0);
                    dtRow[25] = AF;
                    dtRow[26] = Math.Round(AF * AG / 100, 0);
                    dtRow[27] = Z;
                    dtRow[28] = Math.Round(Z * AA / 100, 0);
                    dtRow[29] = AH;
                    dtRow[30] = Math.Round(AH * AI / 100, 0);
                    dtRow[31] = AB;
                    dtRow[32] = Math.Round(AB * AC / 100, 0);
                    dtRow[33] = AJ; dtRow[34] = AK;
                    Huawei.Rows.Add(dtRow);
                    dtRow = Huawei.NewRow();
                }
            }
            return Huawei;
        }
    
        private void DR_2G()
        {
            Opentemplate(template + "\\Bao_cao_CLM_vo_tuyen_2G_Mozambique.xlsx");
            Khaibaosheettemplate("Daily Overview");
            Normal_2G();
            Khaibaosheettemplate("Daily Peak Overview");
            Peak_2G();
            Khaibaosheettemplate("GPRS Overview");
            GPRS_2G();
            Closetemplate(fileout + "\\Bao_cao_CLM_vo_tuyen_2G_Mozambique.xlsx", "2G");  
        }
        private void DR_3G()
        {
            Loaddata();
            Opentemplate(template + "\\Bao_cao_CLM_vo_tuyen_3G_Mozambique.xlsx");
            Khaibaosheettemplate("Daily Province Normal Overview");
            Normal_3G();
            Khaibaosheettemplate("Daily Province Peak Overview");
            Peak_3G();
            Closetemplate(fileout + "\\Bao_cao_CLM_vo_tuyen_3G_Mozambique.xlsx", "3G");
        }       

        private void Normal_2G()
        {
            add_log_file(" ------------------------- Proccessing Normal 2G");
            for (int i = 0; i < dt_ProMap.Rows.Count; i++)
            {
                add_log_file(Title_2 + "[" + dt_ProMap.Rows[i][0].ToString() + "]......");
                float[] Data = CongthucKPI.Normal_2G(NSN, HW, ZTE, Lusr_page, dt_ProMap, i);
                for (int ii = 0; ii < 39; ii++)
                {
                    Rows[14 + i, 5 + ii] = Data[ii];
                }
            }              
        }
        private void Peak_2G()
        {
            add_log_file(" ------------------------- Proccessing Peak 2G");
            for (int i = 0; i < dt_ProMap.Rows.Count; i++)
            {
                add_log_file(Title_2 + "[" + dt_ProMap.Rows[i][0].ToString() + "]......");
                float[] Data = CongthucKPI.Peak_2G(NSN, HW, ZTE, KPI_Peak_2G, dt_ProMap, i);
                for (int ii = 0; ii < 49; ii++)
                {
                    if (ii != 30 && ii != 48)
                    {
                        Rows[14 + i, 13 + ii] = Data[ii];
                    }
                }
            } 
        }
        private void GPRS_2G()
        {

            add_log_file(" ------------------------- Proccessing GPRS 2G");
            for (int i = 0; i < dt_ProMap.Rows.Count; i++)
            {
                add_log_file(Title_2 + "[" + dt_ProMap.Rows[i][0].ToString() + "]......");
                float[] Data = CongthucKPI.GPRS_2G(NSN, HW, ZTE, dt_ProMap, i);
                for (int ii = 0; ii < 18; ii++)
                {
                    Rows[14 + i, 11 + ii] = Data[ii];
                }
            } 
        }

        private void Normal_3G()
        {
            add_log_file(" ------------------------- Proccessing Normal 3G");
            for (int i = 0; i < dt_ProMap.Rows.Count; i++)
            {
                add_log_file(Title_2 + "[" + dt_ProMap.Rows[i][0].ToString() + "]......");
                float[] Data = CongthucKPI.Normal_3G( HW_3G, NSN_3G,Lusr_page,KPI_Normal_3G, dt_ProMap, i);
                for (int ii = 0; ii < 46; ii++)
                {
                    Rows[13 + i, 5 + ii] = Data[ii];
                }
            }   
            
        }
        private void Peak_3G()
        {
            add_log_file(" ------------------------- Proccessing Peak 3G");
            for (int i = 0; i < dt_ProMap.Rows.Count; i++)
            {
                add_log_file(Title_2 + "[" + dt_ProMap.Rows[i][0].ToString() + "]......");
                float[] Data = CongthucKPI.Peak_3G(HW_3G, NSN_3G, Lusr_page, KPI_Peak_3G, dt_ProMap, i);
                for (int ii = 0; ii < 50; ii++)
                {
                    Rows[13 + i, 7 + ii] = Data[ii];
                }
            }   
        }   

        private void Khaibaosheettemplate(string namesheet)
        {
            xlSht = Ewbook.Sheets[namesheet] as Excel.Worksheet;
            xlSht.Name = namesheet;
            Rows = (Excel.Range)xlSht.Cells[1, 1];
        }
        private void Opentemplate(string filetemplate)
        {
            Ewbook = xlApp.Workbooks.Open(filetemplate, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);    
        }
        private void Closetemplate(string filetemplate, string type)
        {
            try
            {
                string[] thoigian = Regex.Split(DateTime.Now.ToString(), @"\W+");
                string date = thoigian[1] + thoigian[0] + thoigian[2] + "_" + thoigian[3] + thoigian[4] + thoigian[5];
                Ewbook.SaveAs(fileout + "\\Bao cao CLM vo tuyen " + type + " Mozambique " + date + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault,
                               null, null, false, false, Excel.XlSaveAsAccessMode.xlExclusive,
                               false, false, false, false, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (System.IO.File.Exists(filetemplate))
                System.IO.File.Delete(filetemplate);
            Ewbook = null;
            xlApp.Quit();
        }
        
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("I Love You");
        }
    }
}
