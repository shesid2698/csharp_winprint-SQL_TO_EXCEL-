using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Xls;
using System.Data;

class SqlConnect
{
    public static DataSet ds = new DataSet();
    private static string Conn
    {
        get
        {
            return @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|BanDongDB.mdf;Integrated Security=True";
        }
    }

    /// <summary>
    /// 測試連線是否成功
    /// </summary>
    /// <returns>成功:success，失敗:fail</returns>
    public static string GetConnect()
    {
        using (SqlConnection cn = new SqlConnection(Conn))
        {
            string test = "fail";
            cn.Open();
            if (cn.State == System.Data.ConnectionState.Open)
            {
                test = "success";
            }
            return test;
        }
    }

    public static void GetTable(DataGridView dgv)
    {
        using (SqlConnection cn = new SqlConnection(Conn))
        {
            cn.Open();
            SqlDataAdapter da=new SqlDataAdapter("select * from Orders",cn);
            da.Fill(ds,"Orders");
            dgv.DataSource= ds.Tables["Orders"];
        }
    }
    public static void SqlToXls()
    {
        using (SqlConnection cn = new SqlConnection(Conn))
        {
            cn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select * from Orders", cn);
            da.Fill(ds, "Orders");
            Workbook wbToStream = new Workbook();
            Worksheet sheet = wbToStream.Worksheets[0];
            sheet["A1"].Text = "餐盒 (便當) 預定/領取登記表";
            sheet["A2"].Text = "訂購日:2022/08/07 上午 10:10:40\n取餐日:2022/08/07 上午 10:10:40\n班級: ";
            sheet["A3"].Text = "編號";
            sheet["B3"].Text = "餐券票號/現金";
            sheet["C3"].Text = "訂購姓名";
            sheet["D3"].Text = "領取姓名";
            sheet["E3"].Text = "主菜選擇";
            sheet["F3"].Text = "備註";
            for (int i = 1; i <= 25; i++)
            {
                sheet.Range[3 + i, 1].Text = $"{i}";
            }
            for (int i = 0; i < ds.Tables["Orders"].Rows.Count; i++)
            {
                sheet.Range[4 + i, 2].Text = ds.Tables["Orders"].Rows[i][1].ToString();
                sheet.Range[4 + i, 3].Text = ds.Tables["Orders"].Rows[i][3].ToString();
                sheet.Range[4 + i, 5].Text = ds.Tables["Orders"].Rows[i][4].ToString();
                sheet.Range[4 + i, 6].Text = ds.Tables["Orders"].Rows[i][5].ToString();
            }
            CellRange xLSContent = sheet["A3:F28"];
            CellRange xLSRanges = sheet["A1:F1"];
            CellRange xLSRanges1 = sheet["A2:F2"];
            xLSRanges.Merge();
            xLSRanges1.Merge();
            XlsStyle.SetXlsHeadStyle(xLSRanges);
            XlsStyle.SetXlsTimeStyle(xLSRanges1);
            XlsStyle.SetXlsBodyStyle(xLSContent);
            sheet.AutoFitColumn(2);
            sheet.AutoFitColumn(6);
            FileStream file_stream = new FileStream("sid.xls", FileMode.Create);
            wbToStream.SaveToStream(file_stream);
            file_stream.Close();
            System.Diagnostics.Process.Start("sid.xls");

        }
    }
}

