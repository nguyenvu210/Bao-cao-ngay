using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;



public class CongthucKPI
{

    public static float[] Rxqual_DL_UL(DataTable dt, string procode)
    {
        float[] Rxqual = new float[2];
        double[] ProDL = new double[8]; double[] ProUL = new double[8];
        double DoL = 0; double UpL = 0;
        double ProAsum = 0; double ProBSum = 0;
        double[] DL = new double[8]; double[] UL = new double[8];
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            string sitename = dt.Rows[i][3].ToString().Trim();
            if (!sitename.Equals(string.Empty) && sitename.Length >= 6)
            {
                string tinh = sitename.Substring(6, 3).Trim().ToUpper();
                if (procode.IndexOf(tinh) != -1)
                {
                    int dem = 6;
                    for (int k = 0; k < 8; k++)
                    {
                        DL[k] = double.Parse(dt.Rows[i][dem].ToString()) + double.Parse(dt.Rows[i][dem + 2].ToString());
                        UL[k] = double.Parse(dt.Rows[i][dem - 1].ToString()) + double.Parse(dt.Rows[i][dem + 1].ToString());
                        dem = dem + 4;
                    }
                }
            }
            for (int k = 0; k < 8; k++)
            {
                ProDL[k] = ProDL[k] + DL[k];
                ProUL[k] = ProUL[k] + UL[k];
            }
        }
        for (int k = 0; k < 8; k++)
        {
            DoL = ProDL[k] * k + DoL;
            UpL = ProUL[k] * k + UpL;
            ProAsum = ProDL[k] + ProAsum;
            ProBSum = ProUL[k] + ProBSum;
        }
        Rxqual[0] = float.Parse((DoL / ProAsum).ToString());
        Rxqual[1] = float.Parse((UpL / ProBSum).ToString());
        return Rxqual;
    }   
    public static float[] get_RxQual(DataTable dt, string procode)
    {
        float[] sum = new float[2];
        for (int j = 0; j < dt.Rows.Count; j++)
        {
            if (dt.Rows[j][1].ToString().ToUpper().Trim().IndexOf(procode.ToUpper()) != -1)
            {
                for (int i = 1; i <= 8; i++)
                {
                    sum[0] += float.Parse(dt.Rows[j][i + 1].ToString()) * i / 100;
                    sum[1] += float.Parse(dt.Rows[j][i + 9].ToString()) * i / 100;
                }
            }
        }
        return sum;
    }
    public static float[] get_average_SQI_VQI_FER(DataTable dt, string procode)
    {
        float[] sum = new float[3];
        for (int j = 0; j < dt.Rows.Count; j++)
        {
            if (dt.Rows[j][1].ToString().ToUpper().Trim().IndexOf(procode.ToUpper()) != -1)
            {
                sum[0] = float.Parse(dt.Rows[j][3].ToString()) + float.Parse(dt.Rows[j][4].ToString()) + float.Parse(dt.Rows[j][12].ToString()) + float.Parse(dt.Rows[j][13].ToString());
                sum[1] = float.Parse(dt.Rows[j][5].ToString()) + float.Parse(dt.Rows[j][6].ToString()) + float.Parse(dt.Rows[j][14].ToString()) + float.Parse(dt.Rows[j][15].ToString());
                sum[2] = float.Parse(dt.Rows[j][7].ToString()) + float.Parse(dt.Rows[j][8].ToString()) + float.Parse(dt.Rows[j][9].ToString()) + float.Parse(dt.Rows[j][10].ToString()) + float.Parse(dt.Rows[j][16].ToString()) + float.Parse(dt.Rows[j][17].ToString()) + float.Parse(dt.Rows[j][18].ToString()) + float.Parse(dt.Rows[j][19].ToString());
            }
        }
        sum[0] = sum[0] / 2;
        sum[1] = sum[1] / 2;
        sum[2] = sum[2] / 2;
        return sum;
    }
    public static float[] Paging_2G_3G(DataTable Paging, DataTable PSR_LUSR, string LAC, string vendor, int boi, string G_2_3, string MSC)
    {
        float k = 0; float l = 0; float G = 0; float H = 0; float E = 0; float F = 0;
        float[] kl = new float[2];
        if (MSC == "NAM")
        {
            string tam = "";
            float GSNA = 0;
            for (int i = 0; i < Paging.Rows.Count; i++)
            {
                string LACID = Paging.Rows[i][3].ToString();
                if (LACID != "" && LACID.Length > 6)
                {
                    LACID = LACID.Substring(LACID.Length - 4, 4).Trim().ToUpper();
                    int myInt = int.Parse(LACID, System.Globalization.NumberStyles.HexNumber);
                    LACID = myInt.ToString().Substring(0, 3);

                    if (LACID.Equals(LAC))
                    {
                        for (int LG = 0; LG < PSR_LUSR.Rows.Count; LG++)
                        {
                            if (Paging.Rows[i][2].ToString() == PSR_LUSR.Rows[LG][0].ToString() && PSR_LUSR.Rows[LG][2].ToString() == G_2_3)
                            {
                                tam = PSR_LUSR.Rows[LG][3].ToString();
                                LG = PSR_LUSR.Rows.Count;
                            }
                        }

                        for (int LH = 0; LH < Paging.Rows.Count; LH++)
                        {

                            if (Paging.Rows[LH][3].ToString().Substring(10, 5).Trim().ToUpper() == tam)
                            {
                                GSNA = float.Parse(Paging.Rows[LH][7 + boi].ToString());
                                LH = Paging.Rows.Count;
                            }
                        }
                        G = float.Parse(Paging.Rows[i][6 + boi].ToString());
                        H = float.Parse(Paging.Rows[i][7 + boi].ToString()) - GSNA;
                        k = H + k;
                        l = G + l;
                    }
                }
            }
            l = l / k * 100;
        }
        if (MSC == "MAC")
        {
            for (int i = 0; i < Paging.Rows.Count; i++)
            {
                string LACID = Paging.Rows[i][2].ToString();
                if (LACID != "" && LACID.Length > 4)
                {
                    LACID = Paging.Rows[i][2].ToString().Substring(0, 3);
                    if (LACID.Equals(LAC))
                    {
                        E = 0; F = 0; G = 0;
                        if (Paging.Rows[i][4].ToString() != "") { E = float.Parse(Paging.Rows[i][4].ToString()); } else { E = 0; }
                        if (Paging.Rows[i][5].ToString() != "") { F = float.Parse(Paging.Rows[i][5].ToString()); } else { F = 0; }
                        if (Paging.Rows[i][6].ToString() != "") { G = float.Parse(Paging.Rows[i][6].ToString()); } else { G = 0; }
                        k = E + F + G + k;
                        l = E + F + l;

                    }
                }
            }
            l = l / k * 100;
        }
        kl[0] = k;
        kl[1] = l;
        return kl;
    }
    public static float[] LUSR_2G_3G(DataTable LUSR, string LAC, string vendor, string MSC)
    {
        float m = 0; float n = 0; float G = 0; float H = 0; float E = 0; float F = 0; float I = 0; float J = 0; float K = 0; float L = 0; float M = 0;
        float[] kl = new float[2];
        if (MSC == "NAM")
        {
            for (int i = 9; i < LUSR.Rows.Count; i++)
            {
                string LACID = LUSR.Rows[i][3].ToString();
                if (LACID != "" && LACID.Length > 6)
                {
                    LACID = LACID.Substring(LACID.Length - 4, 4).Trim().ToUpper();
                    int myInt = int.Parse(LACID, System.Globalization.NumberStyles.HexNumber);
                    LACID = myInt.ToString().Substring(0, 3);

                    if (LACID.Equals(LAC))
                    {
                        E = 0; F = 0;
                        E = float.Parse(LUSR.Rows[i][4].ToString());
                        F = float.Parse(LUSR.Rows[i][5].ToString());
                        m = E + m;
                        n = F + n;
                    }
                }
            }
            n = n / m * 100;
        }
        if (MSC == "MAC")
        {
            for (int i = 0; i < LUSR.Rows.Count; i++)
            {
                string LACID = LUSR.Rows[i][4].ToString();
                if (LACID != "" && LACID.Length > 4)
                {
                    LACID = LUSR.Rows[i][4].ToString().Substring(0, 3);
                    if (LACID.Equals(LAC))
                    {
                        F = 0; G = 0; I = 0; J = 0; K = 0; L = 0; M = 0; H = 0;

                        if (LUSR.Rows[i][5].ToString() != "") { F = float.Parse(LUSR.Rows[i][5].ToString()); } else { F = 0; }
                        if (LUSR.Rows[i][6].ToString() != "") { G = float.Parse(LUSR.Rows[i][6].ToString()); } else { G = 0; }
                        if (LUSR.Rows[i][7].ToString() != "") { H = float.Parse(LUSR.Rows[i][7].ToString()); } else { H = 0; }
                        if (LUSR.Rows[i][8].ToString() != "") { I = float.Parse(LUSR.Rows[i][8].ToString()); } else { I = 0; }
                        if (LUSR.Rows[i][9].ToString() != "") { J = float.Parse(LUSR.Rows[i][9].ToString()); } else { J = 0; }
                        if (LUSR.Rows[i][10].ToString() != "") { K = float.Parse(LUSR.Rows[i][10].ToString()); } else { K = 0; }
                        if (LUSR.Rows[i][11].ToString() != "") { L = float.Parse(LUSR.Rows[i][11].ToString()); } else { L = 0; }
                        if (LUSR.Rows[i][12].ToString() != "") { M = float.Parse(LUSR.Rows[i][12].ToString()); } else { M = 0; }
                        m = F + H + J + L + m;
                        n = G + I + K + M + n;
                    }
                }
            }
            n = n / m * 100;
        }
        kl[0] = m;
        kl[1] = n;
        return kl;
    }

    public static DataTable KPI_Normal_2G(DataSet NSN, DataSet HW, DataSet ZTE)
    {
        DataTable Temp;
        int row;
        int col;
        float[] GET;
        string[] _name = new string[] { "Date", "Vendor", "Pro", "Site", "BSC", "Cell", "SD Attempt", "SD Cong", "SCR", "SD Suc", "SD Drop", "SDR", "TCH Attempt", "TCH Succ", "TASR", "CSSR", "Call Suc", "Call Drop", "CDR", "HO Attempt", "HO Suc", "HOSR", "HI Attempt", "HI Suc", "HISR", "Total Erlang", "Erlang HR", "Traffic SD" };
        DataTable dtNew = new DataTable();
        foreach (string t in _name)
        {
            dtNew.Columns.Add(t);
        }
        DataRow dtRow = dtNew.NewRow();
        #region NOKIA
        Temp = NSN.Tables["Normal"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][3].ToString();
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();

            for (int k = 4; k < col; k++)
            {
                GET[k] = float.Parse(Temp.Rows[i][k].ToString());
            }
            dtRow[0] = Temp.Rows[i][0].ToString();
            dtRow[1] = "Nokia";
            dtRow[2] = tinh;
            dtRow[3] = sitename.Substring(0, 6).Trim().ToUpper();
            dtRow[4] = Temp.Rows[i][1].ToString();
            dtRow[5] = sitename;
            dtRow[6] = GET[30];
            dtRow[7] = GET[10] * GET[30] / 100;
            dtRow[8] = GET[10];
            dtRow[9] = GET[34];
            dtRow[10] = GET[27];
            dtRow[11] = Math.Round(GET[4], 2);
            dtRow[12] = GET[21];
            dtRow[13] = GET[35] + GET[36] + GET[37];
            dtRow[14] = Math.Round(GET[23], 2);
            dtRow[15] = Math.Round(GET[7], 2);
            dtRow[16] = GET[35] + GET[36] + GET[37];
            dtRow[17] = GET[14];
            dtRow[18] = Math.Round(GET[6], 2);
            dtRow[19] = GET[17];
            dtRow[20] = GET[13] * GET[17] / 100;
            dtRow[21] = Math.Round(GET[13], 2);
            dtRow[22] = GET[16];
            dtRow[23] = GET[12] * GET[16] / 100;
            dtRow[24] = Math.Round(GET[12], 2);
            dtRow[25] = Math.Round(GET[25], 2);
            dtRow[26] = Math.Round(GET[31], 2);
            dtRow[27] = Math.Round(GET[28], 2);
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        #region HUAWEI
        Temp = HW.Tables["Normal"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][3].ToString().Substring(6, 7);
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();
            for (int k = 4; k < col; k++)
            {
                if (Temp.Rows[i][k].ToString() != "-")
                    GET[k] = float.Parse(Temp.Rows[i][k].ToString());
                else
                    GET[k] = 0;
            }
            dtRow[0] = Temp.Rows[i][0].ToString();
            dtRow[1] = "Huawei";
            dtRow[2] = tinh;
            dtRow[3] = sitename.Substring(0, 6);
            dtRow[4] = Temp.Rows[i][2].ToString();
            dtRow[5] = sitename;
            dtRow[6] = GET[6];
            dtRow[7] = GET[7];
            dtRow[8] = Math.Round(GET[7] / GET[6] * 100, 2);
            dtRow[9] = GET[10];
            dtRow[10] = GET[9];
            dtRow[11] = Math.Round(GET[9] / GET[10] * 100, 2);
            dtRow[12] = GET[12];
            dtRow[13] = GET[11];
            dtRow[14] = Math.Round(GET[11] / GET[12] * 100, 2);
            dtRow[15] = Math.Round((1 - GET[9] / GET[10]) * GET[11] / GET[12] * 100, 2);
            dtRow[16] = GET[11];
            dtRow[17] = GET[13];
            dtRow[18] = Math.Round(GET[13] / GET[11] * 100, 2);
            dtRow[19] = GET[17];
            dtRow[20] = GET[16];
            dtRow[21] = Math.Round(GET[16] / GET[17] * 100, 2);
            dtRow[22] = GET[15];
            dtRow[23] = Math.Round(GET[14], 2);
            dtRow[24] = Math.Round(GET[14] / GET[15] * 100, 2);
            dtRow[25] = Math.Round(GET[22] * 24, 2);
            dtRow[26] = Math.Round(GET[23] * 24, 2);
            dtRow[27] = Math.Round(GET[21] * 24, 2);
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        #region ZTE
        Temp = ZTE.Tables["Normal"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][11].ToString().Substring(0, 7);
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();
            for (int k = 12; k < col; k++)
            {
                GET[k] = float.Parse(Temp.Rows[i][k].ToString());
            }


            dtRow[0] = Temp.Rows[i][1].ToString();
            dtRow[1] = "ZTE";//Vendor
            dtRow[2] = sitename.Substring(0, 3);//Pro
            dtRow[3] = sitename.Substring(0, 6);//Site
            dtRow[4] = Temp.Rows[i][7].ToString().Substring(0, 6); ;//BSC
            dtRow[5] = sitename;//Cell
            dtRow[6] = GET[21];//SD Attempt
            dtRow[7] = GET[22];//SD Cong
            dtRow[8] = Math.Round(GET[23] * 100, 2);//SCR
            dtRow[9] = GET[52];//SD Suc
            dtRow[10] = GET[28];//SD Drop
            dtRow[11] = Math.Round(GET[29] * 100, 2);//SDR
            dtRow[12] = GET[31];//TCH Attempt
            dtRow[13] = GET[30];//TCH Succ
            dtRow[14] = Math.Round(GET[32] * 100, 2);//TASR
            dtRow[15] = Math.Round(GET[33] * 100, 2);//CSSR
            dtRow[16] = GET[30];//Call Suc
            dtRow[17] = GET[34];//Call Drop
            dtRow[18] = Math.Round(GET[35] * 100, 2);//CDR
            dtRow[19] = GET[39];//HO Attempt
            dtRow[20] = GET[40];//HO Suc
            dtRow[21] = Math.Round(GET[41] * 100, 2);//HOSR
            dtRow[22] = GET[36];//HI Attempt
            dtRow[23] = GET[37];//HI Suc
            dtRow[24] = Math.Round(GET[38] * 100, 2);//HISR
            dtRow[25] = GET[45];//Total Erlang
            dtRow[26] = GET[46];//Erlang HR
            dtRow[27] = GET[44];//Traffic SD
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        for (int i = 0; i < dtNew.Rows.Count; i++)
        {
            for (int j = 0; j < dtNew.Columns.Count; j++)
            {
                if (dtNew.Rows[i][j].ToString() == "NaN" || dtNew.Rows[i][j].ToString() == "Infinity")
                {
                    dtNew.Rows[i][j] = 0;
                }
            }
        }
        return dtNew;
    }
    public static DataTable KPI_Peak_2G(DataSet NSN, DataSet HW, DataSet ZTE)
    {
        DataTable Temp;
        int row;
        int col;
        float[] GET;
        string[] _name = new string[] { "Date", "Vendor", "Pro", "Site", "BSC", "Cell", "TRX", "SDCCH", "TCH", "Trafic offer", "SD Attempt", "SD Cong", "SCR", "SD Suc", "SD Drop", "SDR", "TCH Attempt", "TCH Succ", "TCH Cong", "TCR", "TASR", "CSSR", "Call Suc", "Call Drop", "CDR", "HO Attempt", "HO Suc", "HOSR", "HI Attempt", "HI Suc", "HISR", "Erlang peak", "Erlang HR", "TU (non HR)", "% Traffic HR" };

        DataTable dtNew = new DataTable();
        foreach (string t in _name)
        {
            dtNew.Columns.Add(t);
        }
        DataRow dtRow = dtNew.NewRow();
        #region NOKIA
        Temp = NSN.Tables["Peak"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][3].ToString();
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();

            for (int k = 4; k < col; k++)
            {
                GET[k] = float.Parse(Temp.Rows[i][k].ToString());
            }
            dtRow[0] = Temp.Rows[i][0].ToString();
            dtRow[1] = "Nokia";
            dtRow[2] = tinh;
            dtRow[3] = sitename.Substring(0, 6).Trim().ToUpper();
            dtRow[4] = Temp.Rows[i][1].ToString();
            dtRow[5] = sitename;
            dtRow[6] = Math.Round((1 + GET[46] + (1 + GET[47]) / 8) / 8, 0);
            dtRow[7] = Math.Round(GET[47], 2);
            dtRow[8] = Math.Round(GET[46], 2);
            float tam = float.Parse(ErlangB.Erlang_B(int.Parse(Math.Round(GET[46]).ToString(), 0)).ToString());
            dtRow[9] = Math.Round(tam, 2);
            dtRow[11] = Math.Round(GET[48], 2);
            dtRow[18] = Math.Round(GET[49], 2);
            dtRow[10] = GET[30];
            dtRow[12] = GET[10];
            dtRow[13] = GET[34];
            dtRow[14] = GET[27];
            dtRow[15] = Math.Round(GET[4], 2);
            dtRow[16] = GET[21];
            dtRow[17] = GET[35] + GET[36] + GET[37];
            dtRow[19] = GET[11];
            dtRow[20] = Math.Round(GET[22] / GET[21] * 100, 2);
            dtRow[21] = Math.Round(((1 - Math.Round(GET[4], 2) / 100) * Math.Round(GET[22] / GET[21] * 100, 2)), 2);
            dtRow[22] = GET[35] + GET[36] + GET[37];
            dtRow[23] = GET[14];
            dtRow[24] = Math.Round(GET[6], 2);
            dtRow[25] = GET[17];
            dtRow[26] = Math.Round(GET[13] * GET[17] / 100, 2);
            dtRow[27] = Math.Round(GET[13], 2);
            dtRow[28] = GET[16];
            dtRow[29] = Math.Round(GET[12] * GET[16] / 100, 2);
            dtRow[30] = Math.Round(GET[12], 2);
            dtRow[31] = Math.Round(GET[25], 2);
            dtRow[32] = GET[31];
            dtRow[33] = Math.Round(GET[25] / tam * 100, 2);
            dtRow[34] = Math.Round(GET[31] / GET[25] * 100, 2);
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        #region HUAWEI
        Temp = HW.Tables["Peak"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][4].ToString().Substring(6, 7);
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();

            for (int k = 5; k < col; k++)
            {
                if (Temp.Rows[i][k].ToString() != "-")
                    GET[k] = float.Parse(Temp.Rows[i][k].ToString());
                else
                    GET[k] = 0;
            }
            dtRow[0] = Temp.Rows[i][0].ToString();
            dtRow[1] = "Huawei";
            dtRow[2] = tinh;
            dtRow[3] = sitename.Substring(0, 6).Trim().ToUpper();
            dtRow[4] = Temp.Rows[i][3].ToString().Substring(0, 6);
            dtRow[5] = sitename;
            dtRow[6] = GET[22];
            dtRow[7] = GET[23];
            dtRow[8] = Math.Round(GET[22] * 8 - 1 - GET[23] / 8 - GET[24], 2);
            float tam = float.Parse(ErlangB.Erlang_B(int.Parse((GET[22] * 8 - 1 - GET[23] / 8).ToString())).ToString());
            dtRow[9] = tam;
            dtRow[31] = Math.Round(GET[20], 2);
            dtRow[32] = Math.Round(GET[21], 2);
            dtRow[33] = Math.Round(GET[20] / tam * 100, 2);
            dtRow[34] = Math.Round(GET[21] / GET[20] * 100, 2);
            dtRow[10] = GET[5];
            dtRow[11] = GET[6];
            dtRow[12] = Math.Round(GET[6] / GET[5] * 100, 2);
            dtRow[13] = GET[8];
            dtRow[14] = GET[7];
            dtRow[15] = Math.Round(GET[7] / GET[8] * 100, 2);
            dtRow[16] = GET[10];
            dtRow[17] = Math.Round(GET[9], 2);
            dtRow[18] = GET[15];
            dtRow[19] = Math.Round(GET[15] / GET[10] * 100, 2);
            dtRow[20] = Math.Round(GET[9] / GET[10] * 100, 2);
            dtRow[21] = Math.Round((1 - GET[7] / GET[8]) * GET[9] / GET[10] * 100, 2);
            dtRow[22] = GET[9];
            dtRow[23] = GET[11];
            dtRow[24] = Math.Round(GET[11] / GET[9] * 100, 2);
            dtRow[25] = GET[13];
            dtRow[26] = GET[14];
            dtRow[27] = Math.Round(GET[14] / GET[13] * 100, 2);
            dtRow[28] = GET[16];
            dtRow[29] = GET[12];
            dtRow[30] = Math.Round(GET[12] / GET[16] * 100, 2);
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        #region ZTE
        Temp = ZTE.Tables["Peak"]; row = Temp.Rows.Count; col = Temp.Columns.Count;
        GET = new float[col];
        for (int i = 0; i < row; i++)
        {
            string sitename = Temp.Rows[i][11].ToString();
            string tinh = sitename.Substring(0, 3).Trim().ToUpper();

            for (int k = 12; k < col; k++)
            {
                GET[k] = float.Parse(Temp.Rows[i][k].ToString());
            }
            dtRow[0] = Temp.Rows[i][1].ToString();
            dtRow[1] = "ZTE";//Vendor
            dtRow[2] = tinh;//Pro
            dtRow[3] = sitename.Substring(0, 6);//Site
            dtRow[4] = Temp.Rows[i][7].ToString().Substring(0, 6);//BSC
            dtRow[5] = sitename;//Cell
            dtRow[6] = GET[45];//TRX
            dtRow[7] = GET[47];//SDCCH
            dtRow[8] = GET[46];//TCH
            dtRow[9] = ErlangB.Erlang_B(int.Parse(GET[46].ToString()));//Trafic offer
            dtRow[10] = GET[19];//SD Attempt
            dtRow[11] = GET[20];//SD Cong
            dtRow[12] = Math.Round(GET[21] * 100, 2);//SCR
            dtRow[13] = GET[48];//SD Suc
            dtRow[14] = GET[22];//SD Drop
            dtRow[15] = Math.Round(GET[23] * 100, 2);//SDR
            dtRow[16] = GET[25];//TCH Attempt
            dtRow[17] = GET[24];//TCH Succ
            dtRow[18] = GET[27];//TCH Cong
            dtRow[19] = Math.Round(GET[28] * 100, 2);//TCR
            dtRow[20] = Math.Round(GET[26] * 100 / GET[25] * 100, 2);//TASR
            dtRow[21] = Math.Round(GET[29] * 100, 2);//CSSR
            dtRow[22] = Math.Round(GET[26] * 100, 2);//Call Suc
            dtRow[23] = GET[30];//Call Drop
            dtRow[24] = Math.Round(GET[31] * 100, 2);//CDR
            dtRow[25] = GET[35];//HO Attempt
            dtRow[26] = GET[36];//HO Suc
            dtRow[27] = Math.Round(GET[37] * 100, 2);//HOSR
            dtRow[28] = GET[32];//HI Attempt
            dtRow[29] = GET[33];//HI Suc
            dtRow[30] = Math.Round(GET[34] * 100, 2);//HISR
            dtRow[31] = GET[41];//Erlang peak
            dtRow[32] = GET[42];//Erlang HR
            dtRow[33] = Math.Round(GET[41] / ErlangB.Erlang_B(int.Parse(GET[46].ToString())) * 100, 2);//TU (non HR)
            dtRow[34] = GET[42] / GET[41] * 100;//% Traffic HR
            dtNew.Rows.Add(dtRow);
            dtRow = dtNew.NewRow();
        }
        #endregion
        for (int i = 0; i < dtNew.Rows.Count; i++)
        {
            for (int j = 0; j < dtNew.Columns.Count; j++)
            {
                if (dtNew.Rows[i][j].ToString() == "NaN" || dtNew.Rows[i][j].ToString() == "Infinity")
                {
                    dtNew.Rows[i][j] = 0;
                }
            }
        }
        return dtNew;
    } 
    public static float[] Normal_2G(DataSet NSN, DataSet HW, DataSet ZTE, DataSet Lusr_page ,DataTable Prolist, int ID)
    {
        float[] KQ = new float[40];
        float[] GET;
        string province = Prolist.Rows[ID][0].ToString();
        string Vendor = Prolist.Rows[ID][1].ToString();
        string LAC = Prolist.Rows[ID][2].ToString();
        string MSC = Prolist.Rows[ID][3].ToString();
        int N_site = 0;
        DataTable temp = new DataTable();
   
        if (Vendor == "HW")
        {
            #region Normal
            temp = HW.Tables["Normal"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                sitename = sitename.Substring(6, 7);
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        if (temp.Rows[i][k].ToString() != "-")
                        {
                            GET[k] = float.Parse(temp.Rows[i][k].ToString());
                        }
                        else
                        {
                            GET[k] = 0;
                        }
                    }
                    int check = int.Parse(sitename.Trim().Substring(6, 1).Trim().ToUpper());
                    if (check == 1) { KQ[0]++; }
                    if (check == 7 || check == 4) { KQ[1]++; }
                    if (check < 4) { KQ[2]++; }
                    else { KQ[3]++; }
                    KQ[10] += GET[4];//No. of Random Access Att.
                    KQ[11] += GET[5];// Random Access Succ. Rate (RASR)
                    KQ[12] += GET[6];//No. of SDCCH Ass. Att.
                    KQ[13] += GET[7];//SDCCH Congestion Rate (SCR)
                    KQ[14] += GET[8];//No. of SDCCH Assign Failure
                    KQ[16] += GET[9];//No. of SDCCH Drop
                    KQ[17] += GET[10];//SDCCH Drop Rate (SDR)
                    KQ[18] += GET[11];//No. of TCH Ass. Succ.
                    KQ[19] += GET[12];//TCH Succ. Ass. Rate
                    KQ[21] += GET[13];//No. of TCH Drop
                    KQ[23] += GET[15];//No. of Incoming  HO Att.
                    KQ[24] += GET[14];// Incoming HO Succ. Rate (HISR)
                    KQ[25] += GET[17];//No. of Outgoing  HO Att.
                    KQ[26] += GET[16];// Outgoing HO Succ. Rate (HOSR)
                    KQ[33] += GET[19];//SDCCH Mean Holding Time (SMHT)
                    KQ[34] += GET[20];// TCH Mean Holding Time (TMHT)
                    KQ[35] += GET[21];//Total SDCCH/ A Day
                    KQ[36] += GET[22];//Total Traffic/ A Day
                    KQ[37] += GET[23];//Total Traffic HR/  A Day
                }
            }
            #endregion
            #region Rx_Quality
            temp = HW.Tables["Rx_Quality"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                sitename = sitename.Substring(6, 7);
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    int check = int.Parse(sitename.Trim().Substring(6, 1).Trim().ToUpper());
                    if (check < 4)
                    {
                        KQ[4]++;
                    }
                    else
                    {
                        KQ[5]++;
                    }
                }
            }
            float[] Rxqual = CongthucKPI.Rxqual_DL_UL(temp, province);
            KQ[28] = Rxqual[0] * 100;//RxQual UL
            KQ[29] = Rxqual[1] * 100;//RxQual DL
            KQ[27] = (KQ[28] + KQ[29]) / 2;//RxQual
            #endregion
            #region VQI
            temp = HW.Tables["VQI"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                sitename = sitename.Substring(6, 7);
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    for (int FQ = 5; FQ < 17; FQ++)
                    {
                        KQ[32] += float.Parse(temp.Rows[i][FQ].ToString());
                    }
                    for (int RS = 17; RS < 19; RS++)
                    {
                        KQ[31] += float.Parse(temp.Rows[i][RS].ToString());
                    }
                    for (int TAA = 19; TAA < 27; TAA++)
                    {
                        KQ[30] += float.Parse(temp.Rows[i][TAA].ToString());
                    }
                }
            }          
            #endregion
            #region Paging_2g
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Paging_2g_3g"], Lusr_page.Tables["PSR_LUSR"], LAC, "HW", 0, "2G", MSC);
            KQ[6] = tam[0];//No. of 1st Paging Att.
            KQ[7] = tam[1];//Paging Succ. Rate (PSR)
            #endregion
            #region Lusr_2g
            tam = LUSR_2G_3G(Lusr_page.Tables["LUSR_2g_3g"], LAC, "HW", MSC);
            KQ[8] = tam[0];
            KQ[9] = tam[1];
            #endregion

            KQ[33] = KQ[33] / N_site * 24;//SDCCH Mean Holding Time (SMHT)
            KQ[34] = KQ[34] / N_site * 24;// TCH Mean Holding Time (TMHT)
            KQ[35] = KQ[35] * 24;//Total SDCCH/ A Day
            KQ[36] = KQ[36] * 24;//Total Traffic/ A Day
            KQ[37] = KQ[37] * 24;//Total Traffic HR/  A Day               
            float sum = KQ[30] + KQ[31] + KQ[32]; if (sum == 0) { sum = 1; }
            KQ[32] = KQ[32] / sum * 100;//SQI, VQI, FER Bad
            KQ[31] = KQ[31] / sum * 100;//SQI, VQI, FER Accept
            KQ[30] = KQ[30] / sum * 100;//SQI, VQI, FER Good            
        } 
        if (Vendor == "NSN")
        {
            N_site = 0;
            #region Normal
            temp = NSN.Tables["Normal"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        if (temp.Rows[i][k].ToString() != "-")
                        {
                            GET[k] = float.Parse(temp.Rows[i][k].ToString());
                        }
                        else
                        {
                            GET[k] = 0;
                        }
                    }
                    KQ[13] += GET[49];//SDCCH Congestion Rate (SCR)
                    int check = int.Parse(sitename.Substring(6, 1));

                    if (check == 1)
                    {
                        KQ[0]++;
                    }
                    if (check == 7 || check == 4)
                    {
                        KQ[1]++;
                    }
                    int Su = int.Parse(Math.Round((GET[45] + GET[46] + GET[47] / 8) / 8, 0).ToString());
                    if (check < 4)
                    {
                        KQ[2]++;
                        KQ[4] += Su;
                    }   
                    else
                    {
                        KQ[3]++;
                        KQ[5] += Su;
                    }                   
                    KQ[10] += GET[15];//No. of Random Access Att.
                    KQ[11] += GET[15] * GET[9] / 100;// Random Access Succ. Rate (RASR)
                    KQ[12] += GET[30];//No. of SDCCH Ass. Att.
                    KQ[14] += GET[26];//No. of SDCCH Assign Failure
                    KQ[16] += GET[27];//No. of SDCCH Drop
                    KQ[17] += GET[34];//SDCCH Drop Rate (SDR)
                    KQ[18] += GET[22];//No. of TCH Ass. Succ.
                    KQ[19] += GET[21];//TCH Succ. Ass. Rate
                    KQ[21] += GET[14];//No. of TCH Drop
                    KQ[23] += GET[16];//No. of Incoming  HO Att.
                    KQ[24] += GET[16] * GET[12] / 100;// Incoming HO Succ. Rate (HISR)
                    KQ[25] += GET[17];//No. of Outgoing  HO Att.
                    KQ[26] += GET[17] * GET[13] / 100;// Outgoing HO Succ. Rate (HOSR)
                    KQ[33] += GET[29];//SDCCH Mean Holding Time (SMHT)
                    KQ[34] += GET[24]; // TCH Mean Holding Time (TMHT)
                    KQ[35] += GET[28];//Total SDCCH/ A Day
                    KQ[36] += GET[25];//Total Traffic/ A Day
                    KQ[37] += GET[31];//Total Traffic HR/  A Day                       
                }
            }
            KQ[33] = KQ[33] / N_site;//SDCCH Mean Holding Time (SMHT)
            KQ[34] = KQ[34] / N_site;// TCH Mean Holding Time (TMHT)
            #endregion
            #region  UL_DL_DIS_VQI
            float[] getRxqual = get_RxQual(NSN.Tables["ULDL_ZONEDAY"], province);
            KQ[28] = getRxqual[0];
            KQ[29] = getRxqual[1];
            KQ[27] = (KQ[28] + KQ[29]) / 2;
            float[] Get_VQI = get_average_SQI_VQI_FER(NSN.Tables["DISTRIBUTE_ZONEDAY"], province);
            KQ[30] = Get_VQI[0];
            KQ[31] = Get_VQI[1];
            KQ[32] = Get_VQI[2];
            #endregion
            #region Report_paging
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Report_paging"], Lusr_page.Tables["PSR_LUSR"], LAC, "NSN", 0, "2G", MSC);
            KQ[6] = tam[0];//No. of 1st Paging Att.
            KQ[7] = tam[1];//Paging Succ. Rate (PSR)
            #endregion
            #region LUSR
            tam = LUSR_2G_3G(Lusr_page.Tables["LU_PER_LAC_NETACT3G"], LAC, "NSN", MSC);
            KQ[8] = tam[0];//No. of LU Att.
            KQ[9] = tam[1];//Location Update Succ. Rate (LUSR)
            #endregion
        }
        if (Vendor == "ZTE")
        {
            N_site = 0;
            #region Normal
            temp = ZTE.Tables["Normal"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][11].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 12; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    float Check = float.Parse(temp.Rows[i][10].ToString());
                    if (Check == 1)
                    { KQ[0]++; }
                    if (Check == 4 || Check == 7)
                    { KQ[1]++; }
                    if (Check < 4)
                    { KQ[2]++; KQ[4] += GET[49]; }
                    else
                    { KQ[3]++; KQ[5] += GET[49]; }
                    KQ[10] += GET[19];//No. of Random Access Att.
                    KQ[11] += GET[20] * GET[19] / 100;// Random Access Succ. Rate (RASR)
                    KQ[12] += GET[21];// No. of SDCCH Ass. Att.
                    KQ[13] += GET[22];//SDCCH Congestion Rate (SCR)
                    KQ[14] += GET[26];//No. of SDCCH Assign Failure
                    KQ[16] += GET[28];//No. of SDCCH Drop
                    KQ[17] += GET[52]; //SDCCH Drop Rate (SDR)
                    KQ[18] += GET[30];//No. of TCH Ass. Succ.
                    KQ[19] += GET[31];//TCH Succ. Ass. Rate
                    KQ[21] += GET[34]; //No. of TCH Drop   
                    KQ[23] += GET[36]; //No. of Incoming  HO Att.
                    KQ[24] += GET[37]; // Incoming HO Succ. Rate (HISR)
                    KQ[25] += GET[39];//No. of Outgoing  HO Att.
                    KQ[26] += GET[40];// Outgoing HO Succ. Rate (HOSR)                       
                    KQ[33] += GET[42];//SDCCH Mean Holding Time (SMHT)  
                    KQ[34] += GET[43];// TCH Mean Holding Time (TMHT)
                    KQ[35] += GET[44];//Total SDCCH/ A Day
                    KQ[36] += GET[45];//Total Traffic/ A Day
                    KQ[37] += GET[46];//Total Traffic HR/  A Day                       
                }
            }
            KQ[33] = KQ[33] / N_site;
            KQ[34] = KQ[34] / N_site;
            KQ[11] = KQ[11] * 100;
            #endregion
            #region Paging_lusr
            float[] PAGE = new float[2];
            if (MSC == "MAC")
            {
                PAGE = CongthucKPI.Paging_2G_3G(Lusr_page.Tables["Report_paging"], Lusr_page.Tables["PSR_LUSR"], LAC, "ZTE", 0, "2G", MSC);
            }
            else
            {
                PAGE = CongthucKPI.Paging_2G_3G(Lusr_page.Tables["Paging_2g_3g"], Lusr_page.Tables["PSR_LUSR"], LAC, "ZTE", 0, "2G", MSC);
            }
            KQ[6] = PAGE[0];
            KQ[7] = PAGE[1];
            if (MSC == "MAC")
            {
                PAGE = CongthucKPI.LUSR_2G_3G(Lusr_page.Tables["LU_PER_LAC_NETACT3G"], LAC, "ZTE", MSC);
            }
            else
            {
                PAGE = CongthucKPI.LUSR_2G_3G(Lusr_page.Tables["LUSR_2g_3g"], LAC, "ZTE", MSC);
            }
            KQ[8] = PAGE[0];
            KQ[9] = PAGE[1];
            #endregion
        }
        if (N_site != 0)
        {
            KQ[11] = KQ[11] / KQ[10] * 100;// Random Access Succ. Rate (RASR)
            KQ[13] = KQ[13] / KQ[12] * 100;//SDCCH Congestion Rate (SCR)
            KQ[17] = KQ[16] / KQ[17] * 100;//SDCCH Drop Rate (SDR)
            KQ[15] = KQ[14] / KQ[12] * 100;//SDCCH Assign Failure Rate (SAFR)
            KQ[19] = KQ[18] / KQ[19] * 100;//TCH Succ. Ass. Rate
            KQ[20] = (1 - KQ[17] / 100) * KQ[19];//Call Setup Succ. Rate (CSSR)
            KQ[22] = KQ[21] / KQ[18] * 100;// Call Drop Rate (CDR)
            KQ[24] = KQ[24] / KQ[23] * 100;// Incoming HO Succ. Rate (HISR)
            KQ[26] = KQ[26] / KQ[25] * 100;// Outgoing HO Succ. Rate (HOSR)
            KQ[38] = KQ[37] / KQ[36] * 100;//% Traffic HR/ A Day
        }
        return KQ;
    }
    public static float[] Peak_2G(DataSet NSN, DataSet HW, DataSet ZTE,DataTable KPI_Peak_2G, DataTable Prolist, int ID)
    {
        float[] KQ = new float[49];
        float[] GET;
        string province = Prolist.Rows[ID][0].ToString();
        string Vendor = Prolist.Rows[ID][1].ToString();
        string LAC = Prolist.Rows[ID][2].ToString();
        string MSC = Prolist.Rows[ID][3].ToString();
        int N_site = 0;
        float sum_900_Off = 0;
        float sum_1800_Off = 0;
        float sum_900_RL = 0;
        float sum_1800_RL = 0;
        float count_20 = 0;
        float count_80 = 0;
        DataTable temp = new DataTable();
        #region Huawei
        if (Vendor == "HW")
        {
            temp = HW.Tables["Peak"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][4].ToString();
                sitename = sitename.Substring(6, 7);
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 5; k < temp.Columns.Count; k++)
                    {
                        if (temp.Rows[i][k].ToString() != "-")
                            GET[k] = float.Parse(temp.Rows[i][k].ToString());
                        else
                            GET[k] = 0;
                    }
                    KQ[0] += GET[5];  //No. of SDCCH Ass. Att. Peak Hour
                    KQ[1] += GET[6];//SDCCH Congestion Peak Hour Rate (SCR)
                    KQ[2] += GET[7]; //No. of SDCCH Drop Peak Hour
                    KQ[3] += GET[8];//SDCCH Drop Peak Hour Rate (SDR)
                    KQ[5] += GET[9];//No. of TCH Ass. Succ. Peak Hour
                    KQ[8] += GET[10];//No. of TCH Ass. Att. Peak Hour
                    KQ[9] += GET[15];// TCH Congestion Peak Hour Rate (TCR)
                    KQ[10] += GET[11];//No. of TCH Drop Peak Hour
                    KQ[13] += GET[16];//No. of Incoming  HO Att. Peak Hour
                    KQ[14] += GET[12];// Incoming HO Succ. Peak Hour Rate (HISR)
                    KQ[15] += GET[13];//No. of Outgoing  HO Att. Peak Hour
                    KQ[16] += GET[14];// Outgoing HO Succ. Peak Hour Rate (HOSR)  

                    KQ[17] += GET[17]; //SDCCH Mean Holding Time Peak Hour (SMHT)
                    KQ[18] += GET[18];// TCH Mean Holding Time Peak Hour (TMHT)
                    KQ[19] += GET[19];//SDCCH Traffic Peak Hour
                    KQ[20] += GET[20];// TCH Traffic Peak Hour
                    KQ[21] += GET[21];// TCH Traffic HR Peak Hour                        

                    int check = int.Parse(sitename.Substring(6, 1).ToString());
                    int nu = int.Parse((GET[22] * 8 - Math.Ceiling(GET[23] / 8) - 1 - double.Parse(GET[24].ToString())).ToString());
                    int nu_1_8 = int.Parse(Math.Floor((GET[22] * 8 - Math.Ceiling(GET[23] / 8) - 1 - double.Parse(GET[24].ToString())) * 1.8).ToString());

                    if (check < 4)
                    {
                        KQ[23] += GET[20];//Total Cell TCH Traffic Peak Hour G900
                        KQ[25] += GET[21];//Total Cell TCH Traffic HR Peak Hour G900
                        sum_900_Off = float.Parse(ErlangB.Erlang_B(nu).ToString()) + sum_900_Off;
                        sum_900_RL = GET[20] + sum_900_RL;
                    }
                    else
                    {
                        KQ[24] += GET[20];//Total Cell TCH Traffic Peak Hour G1800
                        KQ[26] += GET[21];//Total Cell TCH Traffic HR Peak Hour G1800
                        sum_1800_Off = float.Parse(ErlangB.Erlang_B(nu).ToString()) + sum_1800_Off;
                        sum_1800_RL = GET[20] + sum_1800_RL;
                    }
                    float b = float.Parse(ErlangB.Erlang_B(nu_1_8).ToString());
                    if (GET[20] / b * 100 < 20) { count_20 = count_20 + 1; }
                    count_80 = count_80 + b;
                    if (GET[22] != 0 && GET[20] != 0)
                    {
                        if (GET[20] / GET[22] < 3)
                        {
                            KQ[31]++;
                        }
                        if (GET[20] / GET[22] > 9)
                        {
                            KQ[32]++;
                        }
                    }
                }
            }
            KQ[28] = sum_900_RL / sum_900_Off * 100;
            KQ[47] = count_20;
            KQ[46] = KQ[20] / count_80 * 100;
            KQ[17] = KQ[17] / N_site;
            KQ[18] = KQ[18] / N_site;
            if (sum_1800_Off != 0)
            {
                KQ[29] = sum_1800_RL / sum_1800_Off * 100;
            }
            KQ[27] = (sum_900_RL + sum_1800_RL) / (sum_900_Off + sum_1800_Off) * 100;
           
        }
        #endregion
        #region Nokia
        if (Vendor == "NSN")
        {
            N_site = 0;
            sum_900_Off = 0; sum_1800_Off = 0;
            sum_900_RL = 0; sum_1800_RL = 0;
            count_20 = 0; count_80 = 0;
            double val = 0; float offer = 0;
            #region Peak
            temp = NSN.Tables["Peak"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[0] += GET[30];//No. of SDCCH Ass. Att. Peak Hour 0
                    KQ[1] += GET[48];
                    KQ[2] += GET[27];//No. of SDCCH Drop Peak Hour 2
                    KQ[3] += GET[34];//No. of SDCCH Drop Peak Hour 2
                    KQ[5] += GET[22];//No. of TCH Ass. Succ. Peak Hour 5
                    KQ[8] += GET[21];//No. of TCH Ass. Att. Peak Hour 8
                    KQ[9] += GET[49];
                    KQ[10] += GET[14];//No. of TCH Drop Peak Hour 10
                    KQ[13] += GET[16];//No. of Incoming  HO Att. Peak Hour 13
                    KQ[14] += GET[12] * GET[16] / 100;// Incoming HO Succ. Peak Hour Rate (HISR) 14
                    KQ[15] += GET[17];//No. of Outgoing  HO Att. Peak Hour 15
                    KQ[16] += GET[13] * GET[17] / 100;// Outgoing HO Succ. Peak Hour Rate (HOSR) 16
                    KQ[17] += GET[29];//SDCCH Mean Holding Time Peak Hour (SMHT) 17
                    KQ[18] += GET[24];// TCH Mean Holding Time Peak Hour (TMHT) 18
                    KQ[19] += GET[28];//SDCCH Traffic Peak Hour 19
                    KQ[20] += GET[25];// TCH Traffic Peak Hour  20
                    KQ[21] += GET[31];// TCH Traffic HR Peak Hour  21

                    int check = int.Parse(sitename.Substring(6, 1));
                    if (check < 4)
                    {
                        KQ[25]++;//Total Cell TCH Traffic HR Peak Hour G900 25
                        sum_900_Off = sum_900_Off + offer;
                        sum_900_RL = GET[25] + sum_900_RL;
                    }
                    else
                    {
                        KQ[26]++;//Total Cell TCH Traffic HR Peak Hour G1800 26
                        sum_1800_Off = sum_1800_Off + offer; ;
                        sum_1800_RL = GET[25] + sum_1800_RL;
                    }
                    val = Math.Ceiling(GET[46]);
                    if (val != 0.0f)
                        offer = float.Parse(ErlangB.Erlang_B(int.Parse(val.ToString())).ToString());
                    if (val > 0)
                    {
                        float b = float.Parse(ErlangB.Erlang_B(int.Parse(Math.Floor(val * 1.8).ToString())).ToString());
                        float a = GET[25] / b * 100;
                        if (a < 20)
                        {
                            count_20 = count_20 + 1;
                        }
                        count_80 = count_80 + b;
                    }
                    float Test = GET[25] / float.Parse(Math.Round(((GET[45] + GET[46] + GET[47] / 8) / 8), 0).ToString());
                    if (Test < 3 && GET[25] != 0)
                    {
                        KQ[31]++;//Số vị trí có Erl/TRx<3 (chỉ tính theo vị trí) 31
                    }
                    if (Test > 9 && GET[25] != 0)
                    {
                        KQ[32]++;//Số vị trí có Erl/TRx<3 (chỉ tính theo vị trí) 32
                    }
                }
            }
            KQ[17] = KQ[17] / N_site;//SDCCH Mean Holding Time Peak Hour (SMHT) 17
            KQ[18] = KQ[18] / N_site;//SDCCH Mean Holding Time Peak Hour (SMHT) 18
            KQ[23] = sum_900_RL;//Total Cell TCH Traffic Peak Hour G900 23
            KQ[24] = sum_1800_RL;//Total Cell TCH Traffic Peak Hour G1800 24
            KQ[27] = KQ[20] / (sum_900_Off + sum_1800_Off) * 100;//TCH Traffic Utilisation Peak Hour (TU) 27
            KQ[28] = sum_900_RL / sum_900_Off * 100;//TCH Traffic Utilisation Peak Hour G900 (TU900) 28
            KQ[29] = sum_1800_RL == 0.0f ? 0.0f : sum_1800_RL / sum_1800_Off * 100;//TCH Traffic Utilisation Peak Hour G1800 (TU1800) 29
            KQ[47] = count_20;//No Cell TU (HR 80%) <20% 47
            KQ[46] = (sum_900_RL + sum_1800_RL) / count_80 * 100;//TU WITH HR 80% 46
            #endregion
        }
        #endregion
        #region ZTE
        if (Vendor == "ZTE")
        {
            N_site = 0;
            sum_900_Off = 0; sum_1800_Off = 0;
            sum_900_RL = 0; sum_1800_RL = 0;
            count_20 = 0; count_80 = 0;
            double val = 0; float offer = 0;
            #region Peak
            temp = ZTE.Tables["Peak"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][11].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 12; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[0] += GET[19];//No. of SDCCH Ass. Att. Peak Hour
                    KQ[1] += GET[20];//SDCCH Congestion Peak Hour Rate (SCR)
                    KQ[2] += GET[22];//No. of SDCCH Drop Peak Hour
                    KQ[3] += GET[48];//SDCCH assignment success
                    KQ[5] += GET[24];//No. of TCH Ass. Succ. Peak Hour
                    KQ[8] += GET[25];//No. of TCH Ass. Att. Peak Hour
                    KQ[9] += GET[27];// TCH Congestion Peak Hour Rate (TCR)
                    KQ[10] += GET[30];//No. of TCH Drop Peak Hour
                    KQ[13] += GET[32];//No. of Incoming  HO Att. Peak Hou
                    KQ[14] += GET[33];// Incoming HO Succ. Peak Hour Rate (HISR)
                    KQ[15] += GET[35];//No. of Outgoing  HO Att. Peak Hour
                    KQ[16] += GET[36];// Outgoing HO Succ. Peak Hour Rate (HOSR)
                    KQ[17] += GET[38];//SDCCH Mean Holding Time Peak Hour (SMHT)                      
                    KQ[18] += GET[39]; // TCH Mean Holding Time Peak Hour (TMHT)     
                    KQ[19] += GET[40];//SDCCH Traffic Peak Hour
                    KQ[20] += GET[41];// TCH Traffic Peak Hour 
                    KQ[21] += GET[42];//  TCH Traffic HR Peak Hour  
                    int K = int.Parse(sitename.Substring(6, 1));
                    if (K < 4)
                    {

                        KQ[23] += GET[41];//Total Cell TCH Traffic Peak Hour G900 
                        KQ[25] += GET[42];//Total Cell TCH Traffic HR Peak Hour G900                            
                        if (GET[46] != 0)
                        {
                            sum_900_Off = sum_900_Off + float.Parse(ErlangB.Erlang_B(int.Parse(GET[46].ToString())).ToString());
                        }
                    }
                    else
                    {
                        KQ[24] += GET[41];// Total Cell TCH Traffic Peak Hour G1800 
                        KQ[26] += GET[42];//Total Cell TCH Traffic HR Peak Hour G1800

                        if (GET[46] != 0)
                        {
                            sum_1800_Off = sum_1800_Off + float.Parse(ErlangB.Erlang_B(int.Parse(GET[46].ToString())).ToString());
                        }
                    }
                    if (GET[44] < 3 && GET[41] != 0)
                    {
                        KQ[31]++; ;//Số vị trí có Erl/TRx<3 (chỉ tính theo vị trí)
                    }
                    if (GET[44] > 9 && GET[41] != 0)
                    {
                        KQ[32]++;//Số vị trí có Erl/TRx>9 (chỉ tính theo vị trí)
                    }
                    if (GET[46] != 0)
                    {
                        val = double.Parse(GET[46].ToString());
                        offer = float.Parse(ErlangB.Erlang_B(int.Parse(Math.Floor(val * 1.8).ToString())).ToString());
                        float TU_Cell = GET[41] / offer * 100;
                        if (TU_Cell < 20)
                        {
                            count_20++;
                        }
                        count_80 = count_80 + offer;
                    }
                }
            }
            KQ[28] = KQ[23] / sum_900_Off * 100;//TCH Traffic Utilisation Peak Hour G900 (TU900)
            KQ[29] = KQ[24] / sum_1800_Off * 100;//TCH Traffic Utilisation Peak Hour G1800 (TU1800)
            KQ[27] = (KQ[23] + KQ[24]) / (sum_1800_Off + sum_900_Off) * 100;//TCH Traffic Utilisation Peak Hour (TU)  
            KQ[17] = KQ[17] / N_site;//SDCCH Mean Holding Time Peak Hour (SMHT)
            KQ[18] = KQ[18] / N_site;// TCH Mean Holding Time Peak Hour (TMHT)   
            KQ[46] = (KQ[23] + KQ[24]) / count_80 * 100;
            KQ[47] = count_20;   
            #endregion
        }          
        #endregion
        #region KPI_Peak_2G
        for (int i = 0; i < KPI_Peak_2G.Rows.Count; i++)
        {
            string tinh =  KPI_Peak_2G.Rows[i][2].ToString();
            GET = new float[KPI_Peak_2G.Rows.Count];
            if (province.IndexOf(tinh) != -1)
            {
                for (int k = 6; k < KPI_Peak_2G.Columns.Count; k++)
                {
                    GET[k] = float.Parse(KPI_Peak_2G.Rows[i][k].ToString());
                }
                if (GET[12] > 0.5 && GET[11] >= 5)
                {
                    KQ[33]++;
                }
                if (GET[12] > 1.25 && GET[11] >= 5)
                {
                    KQ[34]++;
                }
                if (GET[19] > 2 && GET[18] >= 5)
                {
                    KQ[35]++;
                }
                if (GET[19] > 5 && GET[18] >= 5)
                {
                    KQ[36]++;
                }
                if (GET[24] > 2 && GET[23] >= 5)
                {
                    KQ[37]++;
                }
                if (GET[24] > 1 && GET[23] >= 5)
                {
                    KQ[43]++;
                }
                if (GET[24] > 5 && GET[23] >= 5)
                {
                    KQ[38]++;
                }
                if (GET[21] < 97 && GET[16] * (100 - GET[21]) / 100 >= 5)
                {
                    KQ[39]++;
                }
                if (GET[21] < 95 && GET[16] * (100 - GET[21]) / 100 >= 5)
                {
                    KQ[40]++;
                }
                if (GET[21] < 99 && GET[16] * (100 - GET[21]) / 100 >= 5)
                {
                    KQ[44]++;
                }
                if (GET[27] < 95 && GET[25] >= 5)
                {
                    KQ[41]++;
                }
                if (GET[27] < 90 && GET[25] >= 5)
                {
                    KQ[42]++;
                }
            }
        }
        #endregion
        if (N_site != 0)
        {
            KQ[1] = KQ[1] / KQ[0] * 100;//SDCCH Congestion Peak Hour Rate (SCR)
            KQ[3] = KQ[2] / KQ[3] * 100;//SDCCH Drop Peak Hour Rate (SDR)
            KQ[6] = KQ[5] / KQ[8] * 100;//TCH Succ. Ass. Peak Hour Rate
            KQ[7] = (1 - KQ[3] / 100) * KQ[6];//Call Setup Succ. Rate Peak Hour (CSSR)
            KQ[9] = KQ[9] / KQ[8] * 100;// TCH Congestion Peak Hour Rate (TCR)
            KQ[11] = KQ[10] / KQ[5] * 100;// Call Drop Peak Hour Rate (CDR)
            KQ[14] = KQ[14] / KQ[13] * 100;// Incoming HO Succ. Peak Hour Rate (HISR)
            KQ[16] = KQ[16] / KQ[15] * 100;// Outgoing HO Succ. Peak Hour Rate (HOSR)
            KQ[22] = KQ[21] / KQ[20] * 100;//% Traffic HR Peak Hour
        }
        return KQ;
    }   
    public static float[] GPRS_2G(DataSet NSN, DataSet HW, DataSet ZTE, DataTable Prolist, int ID)
    {
        float[] KQ = new float[49];
        float[] GET;
        string province = Prolist.Rows[ID][0].ToString();
        string Vendor = Prolist.Rows[ID][1].ToString();
        string LAC = Prolist.Rows[ID][2].ToString();
        string MSC = Prolist.Rows[ID][3].ToString();
        int N_site = 0;
        float Y2 = 0; float Z2 = 0;
        DataTable temp = new DataTable();
        #region Huawei
        if (Vendor == "HW")
        {
            temp = HW.Tables["GPRS"];
            GET = new float[temp.Columns.Count];
            N_site = 0;
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                sitename = sitename.Substring(6, 7);
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        if (temp.Rows[i][k].ToString() != "OverFlow")
                            GET[k] = float.Parse(temp.Rows[i][k].ToString());
                        else
                            GET[k] = 0;
                    }
                    KQ[0] += GET[4] + GET[5];
                    KQ[1] += GET[6] + GET[7];
                    KQ[3] += GET[8] + GET[9];
                    KQ[4] += GET[10] + GET[11];
                    KQ[6] += GET[12] + GET[13];
                    KQ[8] += GET[14] + GET[15];
                    KQ[10] += GET[4] + GET[5] - GET[16] - GET[17] ;
                    KQ[12] += GET[8] + GET[9] - GET[18] - GET[19] ;
                    KQ[14] += 24 * (
                        GET[22] * 23 + GET[23] * 34 + GET[24] * 40 +
                        GET[25] * 54 + GET[26] * 22 + GET[27] * 28 +
                        GET[28] * 37 + GET[29] * 44 + GET[30] * 56 +
                        GET[31] * 74 + GET[32] * 56 + GET[33] * 68 +
                        GET[34] * 74
                        );
                    Y2 += (20 * (
                        GET[22] + GET[23] + GET[31] +
                        GET[24] + GET[25] + GET[26] +
                        GET[27] + GET[28] + GET[29] +
                        GET[30]
                        ) + 10 * (GET[34] + GET[32] + GET[33]));
                    KQ[15] += 12 * (
                        GET[35] * 23 +
                        GET[36] * 34 + GET[37] * 40 + GET[38] * 54 +
                        GET[39] * 22 + GET[40] * 28 + GET[41] * 37 +
                        GET[42] * 44 + GET[43] * 56 + GET[44] * 74 +
                        GET[45] * 56 + GET[46] * 68 + GET[47] * 74
                        );
                    Z2 += (20 * (
                        GET[44] + GET[35] + GET[36] +
                        GET[37] + GET[38] + GET[39] +
                        GET[40] + GET[41] + GET[42] +
                        GET[43]
                        ) + 10 * (GET[47] + GET[45] + GET[46]));
                    KQ[16] += GET[20] / (1024 * 1024);//aa
                    KQ[17] += GET[21] / (1024 * 1024);//ab
                }
            }         
        }
        #endregion
        #region Nokia
        if (Vendor == "NSN")
        {
            N_site = 0;
            temp = NSN.Tables["GPRS"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[16] += (GET[17] + GET[18]) / 1024;
                    KQ[17] += (GET[19] + GET[20]) / 1024;
                    KQ[14] += 3 * (GET[17] + GET[18]);
                    if(GET[27] == 0)
                    {
                        GET[27] = 1;
                    }
                    if(GET[30] == 0)
                    {
                        GET[30] = 1;
                    }
                    if(GET[28] == 0)
                    {
                        GET[28] = 1;
                    }
                    if(GET[29] == 0)
                    {
                        GET[29] = 1;
                    }
                    if (GET[27] != 0)
                    {
                        Y2 += (GET[17] / GET[27] + GET[18] / GET[30]);
                        Z2 += (GET[20] / GET[28] + GET[19] / GET[29]);
                    }
                    KQ[15] += 3 / 2 * (GET[20] + GET[19]);
                }
            }
             temp = NSN.Tables["GPRS_1"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][3].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    for (int k = 4; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[0] += GET[6] - GET[86] - GET[88];
                    KQ[1] += GET[6];
                    KQ[3] += GET[4] - GET[85] - GET[87];
                    KQ[4] += GET[4] ;
                    KQ[6] += GET[74];
                    KQ[8] += GET[73];
                    KQ[10] += GET[78];
                    KQ[12] += GET[77];
                }
            }
        }
        #endregion
        #region ZTE
        if (Vendor == "ZTE")
        {
            N_site = 0;
            temp = ZTE.Tables["GPRS"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][11].ToString();
                string tinh = sitename.Substring(0, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 12; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[0] += GET[16];
                    KQ[1] += GET[20];
                    KQ[3] += GET[22];
                    KQ[4] += GET[23];
                    KQ[6] += GET[25];
                    KQ[8] += GET[27];
                    KQ[10] += GET[29];
                    KQ[12] += GET[31];
                    KQ[16] += GET[35];
                    KQ[17] += GET[36];
                    KQ[14] += GET[33];
                    KQ[15] += GET[34];
                }
            }
            Y2 = Z2 = N_site;
        }
        #endregion 
      
        if (N_site != 0)
        {
            KQ[2] = KQ[0] / KQ[1] * 100;
            KQ[5] = KQ[3] / KQ[4] * 100;
            KQ[7] = KQ[6] / KQ[1] * 100;
            KQ[9] = KQ[8] / KQ[4] * 100;
            KQ[11] = KQ[10] / KQ[0] * 100;
            KQ[13] = KQ[12] / KQ[3] * 100;
            KQ[14] = KQ[14] / Y2;
            KQ[15] = KQ[15] / Z2;
        }
        return KQ;
    }
   
    #region 3G
    public static float[] Normal_3G(DataSet HW_3G, DataSet NSN_3G, DataSet Lusr_page,DataTable KPI_Normal_3G, DataTable Prolist, int ID)
    {
        float[] KQ = new float[46];
        float[] GET;
        string province = Prolist.Rows[ID][0].ToString();
        string Vendor = Prolist.Rows[ID][1].ToString();
        string LAC = Prolist.Rows[ID][2].ToString();
        string MSC = Prolist.Rows[ID][3].ToString();
        int N_site = 0;
        DataTable temp = new DataTable();

        if (Vendor == "HW")
        {
            #region Normal
            temp = HW_3G.Tables["Normal"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][2].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 6; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }

                    if (int.Parse(sitename.Substring(7, 1)) == 1)
                    {
                        KQ[0] = KQ[0] + 1;
                    }
                    KQ[1]++;
                    KQ[2] += GET[6];//Voice Traffic 2
                    KQ[3] += GET[7];//VP Traffic 3
                    KQ[4] += GET[8];//PS Traffic 4
                    KQ[8] += GET[9];//PS RAB Attempt 8
                    KQ[9] += GET[10];//PS RAB CR 9
                    KQ[10] += GET[11];//CS RAB Attempt 10
                    KQ[11] += GET[12];//CS RAB CR 11
                    KQ[12] += GET[13];//RRC CS Attempt 12
                    KQ[13] += GET[14];//RRC CS SR 13
                    KQ[14] += GET[15];//RRC PS Attempt14
                    KQ[15] += GET[16];//RRC PS SR 15
                    KQ[16] += GET[17];//RAB CS SR 16
                    KQ[17] += GET[18];//RAB PS SR 17
                    KQ[20] += GET[21];//SHO Attempt 20
                    KQ[21] += GET[22];//SHOSR 21
                    KQ[22] += GET[23];//HHO Attempt 22
                    KQ[23] += GET[24];//HHOSR 23
                    KQ[24] += GET[25];//CS Call Attempt 24
                    KQ[25] += GET[26];//CS CDR 25
                    KQ[26] += GET[27]; //CS InRAT HO Attempt 26
                    KQ[27] += GET[28];//CS InRAT HOSR 27
                    KQ[28] += GET[29];//PS Attempt 28
                    KQ[29] += GET[30];//PS CDR 29                   
                    KQ[30] += GET[31]; ;//PS InRAT HO Attempt 30
                    KQ[31] += GET[32];//PS InRAT HOSR 31
                    KQ[32] += GET[33];//HSDPA Throughput  32
                    KQ[33] += GET[34];//HSUPA Throughput 33
                }
            }
            #endregion
            #region Paging_2g_3g
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Paging_2g_3g"], Lusr_page.Tables["PSR_LUSR"], LAC, "HW", 4, "3G", MSC);
            KQ[6] = tam[0];
            KQ[7] = tam[1];
            #endregion
        }        
        if (Vendor == "NSN")
        {
            #region Normal
            temp = NSN_3G.Tables["Normal"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][5].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 7; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    if (int.Parse(sitename.Substring(7, 1)) == 1)
                    {
                        KQ[0]++;
                    }
                    KQ[1]++;
                    KQ[2] += GET[7];//Voice Traffic
                    KQ[3] += GET[8];//VP Traffic	
                    KQ[4] += GET[13];//PS Traffic	
                    KQ[8] += GET[52];//PS RAB Attempt	
                    KQ[9] += GET[40];//PS RAB CR	
                    KQ[10] += GET[49];//CS RAB Attempt
                    KQ[11] += GET[38];//CS RAB CR
                    KQ[12] += GET[43];//RRC CS Attempt
                    KQ[13] += GET[42];//RRC CS SR	
                    KQ[14] += GET[46];//RRC PS Attempt
                    KQ[15] += GET[45];//RRC PS SR	
                    KQ[16] += GET[48];//RAB CS SR	
                    KQ[17] += GET[51];//RAB PS SR		
                    KQ[20] += GET[21];//SHO Attempt
                    KQ[21] += GET[21] * GET[22]/100;//SHOSR
                    KQ[22] += GET[23];//HHO Attempt
                    KQ[23] += GET[24] * GET[25]/100 ;//HHOSR
                    KQ[24] += GET[31];//CS Call Attempt
                    KQ[25] += GET[31] * GET[32]/100 ;//CS CDR
                    KQ[26] += GET[25];//CS InRAT HO Attempt	
                    KQ[27] += GET[25] * GET[26]/100;//CS InRAT HOSR
                    KQ[28] += GET[33];//PS Attempt
                    KQ[29] += GET[33] * GET[34]/100;//PS CDR	
                    KQ[30] += GET[27];//PS InRAT HO Attempt	
                    KQ[31] += GET[27] * GET[28]/100 ;//PS InRAT HOSR	
                    KQ[32] += GET[35];//HSDPA Throughput 	
                    KQ[33] += GET[36];//HSUPA Throughput	
                }
            }
            #endregion
            #region Paging_2g_3g
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Report_paging"], Lusr_page.Tables["PSR_LUSR"], LAC, "NSN", 0, "3G", MSC);
            KQ[6] = tam[0];
            KQ[7] = tam[1];
            #endregion          
        }
        #region KPI_Normal_3G
        float[] kl = KPI_Peak_Normal_3G(KPI_Normal_3G, province);
        KQ[37] = kl[0];
        KQ[38] = kl[1];
        KQ[39] = kl[2];
        KQ[40] = kl[3];
        KQ[41] = kl[4];
        KQ[42] = kl[5];
        KQ[43] = kl[8];
        KQ[44] = kl[9];
        KQ[45] = kl[10];
        #endregion


        if (N_site != 0)
        {
            KQ[15] = KQ[15] / KQ[14] * 100;
            KQ[13] = KQ[13] / KQ[12] * 100;
            KQ[11] = KQ[11] / KQ[10] * 100;
            KQ[9] = KQ[9] / KQ[8] * 100;
            KQ[16] = KQ[16] / KQ[10] * 100;
            KQ[17] = KQ[17] / KQ[8] * 100;
            KQ[23] = KQ[23] / KQ[22] * 100;
            KQ[21] = KQ[21] / KQ[20] * 100;
            KQ[25] = KQ[25] / KQ[24] * 100;
            KQ[27] = KQ[27] / KQ[26] * 100;
            KQ[29] = KQ[29] / KQ[28] * 100;
            KQ[31] = KQ[31] / KQ[30] * 100;
            KQ[18] = KQ[13] * KQ[16] / 100;
            KQ[19] = KQ[15] * KQ[17] / 100;
            KQ[5] = KQ[5] / KQ[30] * 100;
            KQ[32] = KQ[32] / N_site;
            KQ[33] = KQ[33] / N_site;
        }
        return KQ;
    }
    public static float[] Peak_3G(DataSet HW_3G, DataSet NSN_3G, DataSet Lusr_page, DataTable KPI_Peak_3G, DataTable Prolist, int ID)
    {
        float[] KQ = new float[50];
        float[] GET;
        string province = Prolist.Rows[ID][0].ToString();
        string Vendor = Prolist.Rows[ID][1].ToString();
        string LAC = Prolist.Rows[ID][2].ToString();
        string MSC = Prolist.Rows[ID][3].ToString();
        int N_site = 0;
        DataTable temp = new DataTable();

        if (Vendor == "HW")
        {
            #region PeakCS
            temp = HW_3G.Tables["Peak_CS"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
               
                string sitename = temp.Rows[i][2].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 6; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[0] += GET[6];
                    KQ[1] += GET[7];
                    KQ[8] += GET[8];
                    KQ[9] += GET[9];
                    KQ[10] += GET[11];
                    KQ[11] += GET[12];
                    KQ[14] += GET[10];
                    KQ[22] += GET[14];
                    KQ[23] += GET[15];
                    KQ[24] += GET[16];
                    KQ[25] += GET[17];              
                }
            }
            #endregion
            #region PeakPS
            temp = HW_3G.Tables["Peak_PS"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][2].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    for (int k = 6; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }

                    KQ[2] += GET[6];
                    KQ[6] += GET[7];
                    KQ[7] += GET[8];
                    KQ[12] += GET[10];
                    KQ[13] += GET[11];
                    KQ[15] += GET[9];
                    KQ[26] += GET[13];
                    KQ[27] += GET[14];
                    KQ[28] += GET[15];
                    KQ[29] += GET[16];
                    KQ[30] += GET[17];
                    KQ[31] += GET[18]; 
                }
            }
            #endregion
            #region Peak
            temp = HW_3G.Tables["Peak"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][2].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    for (int k = 6; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }
                    KQ[18] += GET[21];
                    KQ[19] += GET[22];
                    KQ[20] += GET[23];
                    KQ[21] += GET[24]; 
                }
            }
            #endregion
            #region Paging_2g_3g
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Paging_2g_3g"], Lusr_page.Tables["PSR_LUSR"], LAC, "HW", 4, "3G", MSC);
            KQ[4] = tam[0];
            KQ[5] = tam[1];
            #endregion
        }
        if (Vendor == "NSN")
        {
            #region Peak
            temp = NSN_3G.Tables["Peak"];
            GET = new float[temp.Columns.Count];
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                string sitename = temp.Rows[i][5].ToString();
                string tinh = sitename.Substring(1, 3).Trim().ToUpper();
                if (province.IndexOf(tinh) != -1)
                {
                    N_site++;
                    for (int k = 7; k < temp.Columns.Count; k++)
                    {
                        GET[k] = float.Parse(temp.Rows[i][k].ToString());
                    }                 
                    KQ[0] += GET[7];//Voice Traffic
                    KQ[1] += GET[8];//VP Traffic	
                    KQ[2] += GET[13];//PS Traffic	
                    KQ[6] += GET[52];//PS RAB Attempt	
                    KQ[7] += GET[40];//PS RAB CR	
                    KQ[8] += GET[49];//CS RAB Attempt
                    KQ[9] += GET[38];//CS RAB CR
                    KQ[10] += GET[43];//RRC CS Attempt
                    KQ[11] += GET[42];//RRC CS SR	
                    KQ[12] += GET[46];//RRC PS Attempt
                    KQ[13] += GET[45];//RRC PS SR	
                    KQ[14] += GET[48];//RAB CS SR	
                    KQ[15] += GET[51];//RAB PS SR		
                    KQ[18] += GET[21];//SHO Attempt
                    KQ[19] += GET[21] * GET[22] / 100;//SHOSR
                    KQ[20] += GET[23];//HHO Attempt
                    KQ[21] += GET[24] * GET[25] / 100;//HHOSR
                    KQ[22] += GET[31];//CS Call Attempt
                    KQ[23] += GET[31] * GET[32] / 100;//CS CDR
                    KQ[24] += GET[25];//CS InRAT HO Attempt	
                    KQ[25] += GET[25] * GET[26] / 100;//CS InRAT HOSR
                    KQ[26] += GET[33];//PS Attempt
                    KQ[27] += GET[33] * GET[34] / 100;//PS CDR	
                    KQ[28] += GET[27];//PS InRAT HO Attempt	
                    KQ[29] += GET[27] * GET[28] / 100;//PS InRAT HOSR	
                    KQ[30] += GET[35];//HSDPA Throughput 	
                    KQ[31] += GET[36];//HSUPA Throughput	
                }
            }
            #endregion
            #region Paging_2g_3g
            float[] tam = Paging_2G_3G(Lusr_page.Tables["Report_paging"], Lusr_page.Tables["PSR_LUSR"], LAC, "NSN", 0, "3G", MSC);
            KQ[4] = tam[0];
            KQ[5] = tam[1];
            #endregion
        }
        #region KPI_Peak_3G
        float[] kl = KPI_Peak_Normal_3G(KPI_Peak_3G, province);
        KQ[42] = kl[0];
        KQ[43] = kl[1];
        KQ[44] = kl[2];
        KQ[45] = kl[3];
        KQ[46] = kl[4];
        KQ[47] = kl[5];
        KQ[48] = kl[6];
        KQ[49] = kl[7];
        #endregion
        if (N_site != 0)
        {
            KQ[13] = KQ[13] / KQ[12] * 100;
            KQ[11] = KQ[11] / KQ[10] * 100;
            KQ[9] = KQ[9] / KQ[8] * 100;
            KQ[7] = KQ[7] / KQ[6] * 100;
            KQ[14] = KQ[14] / KQ[8] * 100;
            KQ[15] = KQ[15] / KQ[6] * 100;
            KQ[21] = KQ[21] / KQ[20] * 100;
            KQ[19] = KQ[19] / KQ[18] * 100;
            KQ[23] = KQ[23] / KQ[22] * 100;
            KQ[25] = KQ[25] / KQ[24] * 100;
            KQ[27] = KQ[27] / KQ[26] * 100;
            KQ[29] = KQ[29] / KQ[28] * 100;
            KQ[16] = KQ[11] * KQ[14] / 100;
            KQ[17] = KQ[13] * KQ[15] / 100;
            KQ[3] = KQ[3] / KQ[28] * 100;
            KQ[30] = KQ[30] / N_site;
            KQ[31] = KQ[31] / N_site;        }
       
        return KQ;
    }
    #endregion
    public static float[] KPI_Peak_Normal_3G(DataTable KPIG, string procode)
    {
        float bb = 0; float bc = 0; float aw = 0; float ay = 0;
        float ax = 0; float az = 0; float ba = 0; float bd = 0; float a3 = 0; float a1 = 0; float a2 = 0;
        float[] kl = new float[11];
        for (int ii = 0; ii < KPIG.Rows.Count; ii++)
        {
            string sitename = KPIG.Rows[ii][2].ToString().Substring(1, 3).ToUpper().ToString();

            if (sitename.Equals(procode))
            {
                float J = 0; float K = 0; float L = 0; float M = 0; float N = 0; float P = 0; float S = 0;
                float Z = 0; float AD = 0; float AE = 0; float AH = 0; float AI = 0;
                float AA = 0; float T = 0; float U = 0; float R = 0; float O = 0; float Q = 0;
                if (KPIG.Rows[ii][12].ToString() != "NaN" && KPIG.Rows[ii][12].ToString() != "") { M = float.Parse(KPIG.Rows[ii][12].ToString()); } else { M = 0; };
                if (KPIG.Rows[ii][10].ToString() != "NaN" && KPIG.Rows[ii][10].ToString() != "") { K = float.Parse(KPIG.Rows[ii][10].ToString()); } else { K = 0; };
                if (KPIG.Rows[ii][18].ToString() != "NaN" && KPIG.Rows[ii][18].ToString() != "") { S = float.Parse(KPIG.Rows[ii][18].ToString()); } else { S = 0; };
                if (KPIG.Rows[ii][9].ToString() != "NaN" && KPIG.Rows[ii][9].ToString() != "") { J = float.Parse(KPIG.Rows[ii][9].ToString()); } else { J = 0; };
                if (KPIG.Rows[ii][19].ToString() != "NaN" && KPIG.Rows[ii][19].ToString() != "") { T = float.Parse(KPIG.Rows[ii][19].ToString()); } else { T = 0; };
                if (KPIG.Rows[ii][13].ToString() != "NaN" && KPIG.Rows[ii][13].ToString() != "") { N = float.Parse(KPIG.Rows[ii][13].ToString()); } else { N = 0; };
                if (KPIG.Rows[ii][20].ToString() != "NaN" && KPIG.Rows[ii][20].ToString() != "") { U = float.Parse(KPIG.Rows[ii][20].ToString()); } else { U = 0; };
                if (KPIG.Rows[ii][15].ToString() != "NaN" && KPIG.Rows[ii][15].ToString() != "") { P = float.Parse(KPIG.Rows[ii][15].ToString()); } else { P = 0; };
                if (KPIG.Rows[ii][26].ToString() != "NaN" && KPIG.Rows[ii][26].ToString() != "") { AA = float.Parse(KPIG.Rows[ii][26].ToString()); } else { AA = 0; };
                if (KPIG.Rows[ii][25].ToString() != "NaN" && KPIG.Rows[ii][25].ToString() != "") { Z = float.Parse(KPIG.Rows[ii][25].ToString()); } else { Z = 0; };
                if (KPIG.Rows[ii][30].ToString() != "NaN" && KPIG.Rows[ii][30].ToString() != "") { AE = float.Parse(KPIG.Rows[ii][30].ToString()); } else { AE = 0; };
                if (KPIG.Rows[ii][29].ToString() != "NaN" && KPIG.Rows[ii][29].ToString() != "") { AD = float.Parse(KPIG.Rows[ii][29].ToString()); } else { AD = 0; };
                if (KPIG.Rows[ii][33].ToString() != "NaN" && KPIG.Rows[ii][33].ToString() != "") { AH = float.Parse(KPIG.Rows[ii][33].ToString()); } else { AH = 0; };
                if (KPIG.Rows[ii][34].ToString() != "NaN" && KPIG.Rows[ii][34].ToString() != "") { AI = float.Parse(KPIG.Rows[ii][34].ToString()); } else { AI = 0; };
                if (KPIG.Rows[ii][18].ToString() != "NaN" && KPIG.Rows[ii][17].ToString() != "") { R = float.Parse(KPIG.Rows[ii][17].ToString()); } else { R = 0; };
                if (KPIG.Rows[ii][15].ToString() != "NaN" && KPIG.Rows[ii][14].ToString() != "") { O = float.Parse(KPIG.Rows[ii][14].ToString()); } else { O = 0; };
                if (KPIG.Rows[ii][16].ToString() != "NaN" && KPIG.Rows[ii][16].ToString() != "") { Q = float.Parse(KPIG.Rows[ii][16].ToString()); } else { Q = 0; };
                if (M / L * 100 > 5 && M >= 5)
                {
                    aw = aw + 1;
                }
                if (K / J * 100 > 5 && K >= 5)
                {
                    ax = ax + 1;
                }
                if (T < 95 && (L - R + N - O) >= 5)// N * (100 - T) / 100 >= 5)
                {
                    ay = ay + 1;    //ar                 
                }
                if (T < 99 && (L - R + N - O) >= 5)// N * (100 - T) / 100 >= 5)
                {
                    a2 = a2 + 1;    //ar                 
                }
                if (U < 95 && (J - S + P - Q) >= 5)//P * (100 - U) / 100 >= 5)
                {
                    az = az + 1;    //aw
                }
                if (U < 99 && (J - S + P - Q) >= 5)//P * (100 - U) / 100 >= 5)
                {
                    a3 = a3 + 1;    //ax
                }
                if (AA / Z * 100 > 5 && AA >= 5)
                {
                    ba = ba + 1;
                }
                if (AA / Z * 100 > 1 && AA >= 5)//////////
                {
                    a1 = a1 + 1;//av
                }
                if (AE / AD * 100 > 5 && AE >= 5)
                {
                    bb = bb + 1;
                }
                if ((AH * 1000) < 800)
                {
                    bc = bc + 1;
                }
                if ((AI * 1000) < 256)
                {
                    bd = bd + 1;
                }
            }
        }
        kl[0] = aw;//ap
        kl[1] = ax;//aq
        kl[2] = ay;//ar
        kl[3] = az;//as
        kl[4] = bb;//at
        kl[5] = ba;//au
        kl[6] = bc;
        kl[7] = bd;
        kl[8] = a1;//av
        kl[9] = a2;//aw
        kl[10] = a3;//ax

        return kl;
    }
}


