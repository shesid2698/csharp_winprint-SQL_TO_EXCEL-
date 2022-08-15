using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;

namespace winprint
{
    public partial class SQL_EXCEL : Form
    {
        public SQL_EXCEL()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnect.SqlToXls();
            //A: Dynamically create Excel file and save it to stream
            //Workbook wbToStream = new Workbook();
            //Worksheet sheet = wbToStream.Worksheets[0];
            ////sheet["C1"].Text = "The sample demonstrates how to save an Excel workbook to stream.";
            ////sheet["B1"].Text = "111111111111111111111111";
            ////sheet["B2"].Text = "2222222222222222";
            //sheet["A1"].Text = "餐盒(便當) 預定 / 領取登記表";
            //sheet["A1"].Style.Font.FontName = "標楷體";
            //sheet["A1"].Style.Font.Size = 16;
            //sheet["A1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
            //sheet["A1"].AutoFitRows();
            //sheet["A1:F1"].Merge();
            ////CellStyle style = wbToStream.Styles.Add("sty");
            //CellRange xLSRanges = sheet.Range["A1:F1"];
            //xLSRanges.Style.Color = Color.Blue;


            ////sheet.Rows[1].Merge
            ////sheet["b1:g1"].Merge();


            //sheet.AutoFitColumn(2);
            ////sheet.AutoFitColumn(3);
            ////第幾欄自動調整
            //FileStream file_stream = new FileStream("sid.xls", FileMode.Create);
            //wbToStream.SaveToStream(file_stream);
            //file_stream.Close();
            //System.Diagnostics.Process.Start("sid.xls");

            //B. Load Excel file from stream
            //Workbook wbFromStream = new Workbook();
            //FileStream fileStream = File.OpenRead("sample.xls");
            //fileStream.Seek(0, SeekOrigin.Begin);
            //wbFromStream.LoadFromStream(fileStream);
            //wbFromStream.SaveToFile("From_stream.xls", ExcelVersion.Version97to2003);
            //fileStream.Dispose();
            //System.Diagnostics.Process.Start("From_stream.xls");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SqlConnect.GetTable(dgv);
        }

        private void btn_print_Click(object sender, EventArgs e)
        {

            Workbook workbook = new Workbook();
            workbook.LoadFromFile("C:\\C#\\winprint\\bin\\Debug\\sid.xls");
            PrintDialog dialog = new PrintDialog();
            dialog.AllowPrintToFile = true;
            dialog.AllowCurrentPage = true;
            dialog.AllowSomePages = true;
            dialog.AllowSelection = true;
            dialog.UseEXDialog = true;
            dialog.PrinterSettings.Duplex = Duplex.Simplex;
            dialog.PrinterSettings.FromPage = 0;
            dialog.PrinterSettings.ToPage = 8;
            dialog.PrinterSettings.PrintRange = PrintRange.SomePages;
            workbook.PrintDialog = dialog;
            PrintDocument pd = workbook.PrintDocument;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                pd.Print();
            }
        }
    }
}



