using app;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
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
        public static Dictionary<string, string> marketNamesMax7DayZarar = new Dictionary<string, string>();
        public static Dictionary<string, string> marketNamesId = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnOnlineCheck_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            try
            {
                var request = WebRequest.CreateHttp("http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0");
                var response = request.GetResponse();
                byte[] tempBytes = new byte[4096];


                using (var stream = new MemoryStream())
                {
                    using (var newstream = response.GetResponseStream())
                    {
                        using (GZipStream zipStream = new GZipStream(newstream, CompressionMode.Decompress))
                        {
                            //zipStream.CopyTo()
                            int i;
                            while ((i = zipStream.Read(tempBytes, 0, tempBytes.Length)) != 0)
                            {
                                stream.Write(tempBytes, 0, i);
                            }
                        }


                        //Stream stream = new MemoryStream(tempBytes);
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
                                    var marketName = mainSheet.Cells[$"A{row}"].Value.ToString();
                                    if (string.IsNullOrEmpty(marketName) || marketName.Any(c => char.IsDigit(c)) || marketName.EndsWith('ح'.ToString()))
                                        continue;
                                    var localResultSheet = localExcel.Workbook.Worksheets.Where(n => n.Name == marketName).FirstOrDefault();
                                    if (localResultSheet == null)
                                    {
                                        continue;
                                    }

                                    int localResultRowCount = localResultSheet.Dimension.Rows;
                                    //! درصد لحظه ای
                                    float currentDarsad = float.Parse(mainSheet.Cells[$"J{row}"].Value.ToString());
                                    float currentPayani = float.Parse(mainSheet.Cells[$"M{row}"].Value.ToString());
                                    float lastDayDarsad = float.Parse((localResultSheet.Cells[$"L{localResultRowCount}"].Value ?? "0").ToString());
                                    if (lastDayDarsad < -1 && currentDarsad > 3 && currentPayani > 1)
                                    {
                                        textBox1.Text += $"{marketName}     {currentDarsad}     {currentPayani}     دیروز:{lastDayDarsad}" + Environment.NewLine;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetCurrentDate()
        {
            var d = DateTime.Now.AddDays(0);
            PersianCalendar pc = new PersianCalendar();
            return string.Format("{0}/{1}/{2}", pc.GetYear(d), pc.GetMonth(d).ToString().PadLeft(2, '0'), pc.GetDayOfMonth(d).ToString().PadLeft(2, '0'));
        }

        private void btnOnlineCheckMaxZarar_Click(object sender, EventArgs e)
        {
            try
            {
                string exResult = @"N:\Bourse\marketResult.xlsx";
                ExcelObjectCompare excelObjectCompare = new ExcelObjectCompare();
                FileInfo fileResult = new FileInfo(exResult);
                marketNamesMax7DayZarar = new Dictionary<string, string>();

                //var handler = new HttpClientHandler();
                //handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                var client = new WebClient();
                client.Headers[HttpRequestHeader.AcceptEncoding] = "gzip";
                client.Encoding = Encoding.Unicode;

                using (ExcelPackage result = new ExcelPackage(fileResult))
                {
                    //var orderedMaxZarar3Day = result.Workbook.Worksheets
                    //                                .Where(n => n.Name != "Charts")
                    //                                .Where(n => n.Cells[$"BC{n.Dimension.Rows}"] != null && n.Cells[$"BC{n.Dimension.Rows}"].Value != null && (double)n.Cells[$"BC{n.Dimension.Rows}"].Value < 3000)
                    //                                .OrderBy(n => n.Cells[$"BC{n.Dimension.Rows}"], excelObjectCompare)
                    //                                .ToList().Take(20);



                    var orderedMaxZarar7Day = result.Workbook.Worksheets
                                        .Where(n => n.Name != "Charts")
                                        .Where(n => n.Cells[$"BD{n.Dimension.Rows}"] != null && n.Cells[$"BD{n.Dimension.Rows}"].Value != null && (double)n.Cells[$"BD{n.Dimension.Rows}"].Value < 3000)
                                        .OrderBy(n => n.Cells[$"BD{n.Dimension.Rows}"], excelObjectCompare)
                                        .ToList().Take(20);
                    orderedMaxZarar7Day.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = max - 7;
                        var data = new List<object>();
                        //! L ==> درصد پایانی دیروز

                        client.Headers[HttpRequestHeader.AcceptEncoding] = "gzip, deflate";
                        client.Encoding = Encoding.Unicode;
                        string searchMarketURL = $"http://tsetmc.com/tsev2/data/search.aspx?skey={n.Name}";
                        var findNameResult = client.DownloadString(searchMarketURL);
                        var findedItems = findNameResult.Split(';');
                        var marketId = findedItems.Where(m => m.Split(',')[0] == n.Name).Select(m => m.Split(',')[2]).First().ToString();

                        marketNamesMax7DayZarar[n.Name] = n.Cells[$"L{max}"].Value.ToString();
                        marketNamesId[n.Name] = marketId;
                    });
                }

                System.Threading.Tasks.Task.Factory.StartNew(AtMomentCheck);
            }
            catch (Exception ex)
            {

            }
        }

        static void AtMomentCheck()
        {
            foreach (var marketName in marketNamesMax7DayZarar.Keys)
            {
                var lastDayPayanyDarsad = float.Parse(marketNamesMax7DayZarar[marketName]);
                var marketId = marketNamesId[marketName];

                var client = new WebClient();
                string marketOnlineInfoUrl = $"http://www.tsetmc.com/tsev2/data/instinfodata.aspx?i={marketId}&c=41+";
                var marketOnlineInfo = client.DownloadString(marketOnlineInfoUrl);
                var marketOnlineSplitInfo = marketOnlineInfo.Split(';');
                var hogogiInfo = marketOnlineSplitInfo[4].Split(',');
                if (hogogiInfo.Length == 10)
                {
                    var hogogiSeal = hogogiInfo[4];
                    var hogogiBay = hogogiInfo[1];
                    var hagigiBay = hogogiInfo[0];
                    var hagigiSeal = hogogiInfo[3];
                    var mainInfo = marketOnlineSplitInfo[0].Split(',');
                    var hagmKol = mainInfo[9];
                    var mablagKol = mainInfo[10];
                }
                else
                {


                }
            }
        }
    }
}
