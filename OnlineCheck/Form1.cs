using app;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OnlineCheck
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void btnOnlineCheck_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            try
            {
                string getTsemcExcelURL = "http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0";
                var handler = new HttpClientHandler();
                handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                var client = new HttpClient(handler);

                var downloadedExcel = await client.GetByteArrayAsync(getTsemcExcelURL);
                Stream stream = new MemoryStream(downloadedExcel);

                using (ExcelPackage marketExcel = new ExcelPackage(stream))
                {
                    var mainSheet = marketExcel.Workbook.Worksheets.Where(n => n.Name == "دیده بان بازار").FirstOrDefault();
                    string exResult = "marketResult.xlsx";
                    FileInfo fileResult = new FileInfo(exResult);
                    using (ExcelPackage localExcel = new ExcelPackage(fileResult))
                    {
                        int rowCount = mainSheet.Dimension.Rows;
                        int ColCount = mainSheet.Dimension.Columns;
                        for (int row = 4; row <= rowCount; row++)
                        {
                            var marketName = mainSheet.Cells[row, 1].Value.ToString();
                            if (string.IsNullOrEmpty(marketName) || marketName.Any(c => char.IsDigit(c)) || marketName.EndsWith('ح'.ToString()))
                                continue;
                            var localResultSheet = localExcel.Workbook.Worksheets.Where(n => n.Name == marketName).FirstOrDefault();
                            if (localResultSheet == null)
                            {
                                continue;
                            }

                            int localResultRowCount = mainSheet.Dimension.Rows;
                            //! درصد لحظه ای
                            float currentDarsad = float.Parse(mainSheet.Cells[$"J{row}"].Value.ToString());
                            float currentPayani = float.Parse(mainSheet.Cells[$"J{row}"].Value.ToString());
                            float lastDayDarsad = float.Parse(localResultSheet.Cells[$"L{localResultRowCount}"].Value.ToString());
                            if(lastDayDarsad<-1 && currentDarsad>3 && currentPayani>1)
                            {
                                textBox1.Text += $"{marketName}     {currentDarsad}     {currentPayani}     دیروز:{lastDayDarsad}";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("chart", ex);
            }
        }

        private string GetCurrentDate()
        {
            var d = DateTime.Now.AddDays(0);
            PersianCalendar pc = new PersianCalendar();
            return string.Format("{0}/{1}/{2}", pc.GetYear(d), pc.GetMonth(d).ToString().PadLeft(2, '0'), pc.GetDayOfMonth(d).ToString().PadLeft(2, '0'));
        }
    }
}
