using app;
using Fleck;
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
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
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

                //System.Threading.Tasks.Task.Factory.StartNew(AtMomentCheck);
            }
            catch (Exception ex)
            {

            }
        }

        private async Task<string> AtMomentCheck(string marketid, string wantedStatus)
        {
            //foreach (var marketName in marketNamesMax7DayZarar.Keys)
            //{
            //    var lastDayPayanyDarsad = float.Parse(marketNamesMax7DayZarar[marketName]);
            //    var marketId = marketNamesId[marketName];

            var handler = new HttpClientHandler();
            handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            var client = new HttpClient(handler);
            //var client = new WebClient();
            string marketOnlineInfoUrl = $"http://www.tsetmc.com/tsev2/data/instinfodata.aspx?i={marketid}&c=41+";
            var marketOnlineInfo = await client.GetStringAsync(marketOnlineInfoUrl);
            var marketOnlineSplitInfo = marketOnlineInfo.Split(';');
            var currentInfos = marketOnlineSplitInfo[0].Split(',');
            if (string.IsNullOrEmpty(wantedStatus))
            {
                //! AR مجاز-محفوظ -- بعد از مچ شدن همون موقع شروع به ادامه خرید میکنه
                if (currentInfos[1].Trim() != "IS" && currentInfos[1].Trim() != "I" && currentInfos[1].Trim() != "AR")
                {
                    sockets.ForEach(n => n.Send("By"));
                    MessageBox.Show("End");
                    return "End";
                }
                textBox1.Text += currentInfos[1].Trim() + ":" + DateTime.Now.ToString("mm:ss") + ",";
            }
            else if (wantedStatus == "CheckRizeshSaf")
            {
                var karidoforosh = marketOnlineSplitInfo[2].Split('@');
                var forosh = karidoforosh[4];
                //! یک میلیون
                //! باید از ورودی خوانده شود
                if (int.Parse(forosh) < 1000000)
                {
                    sockets.ForEach(n => n.Send("OneBy"));
                    MessageBox.Show("End");
                    return "End";
                }
                textBox1.Text += DateTime.Now.ToString("mm:ss") + "==> " + forosh + ", ";
            }


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
            //}
            return null;
        }

        private async void btnOneMaket_Click(object sender, EventArgs e)
        {
            var res = string.Empty;
            do
            {
                try
                {
                    res = await AtMomentCheck(txtMarketId.Text, "");
                    System.Threading.Thread.Sleep(1000);
                }
                catch (Exception)
                {
                    sockets.ForEach(n => n.Send("ByOne"));
                    System.Threading.Thread.Sleep(100);
                }
            } while (res != "End");
            //sockets.ForEach(n => n.Send("By"));
        }



        List<IWebSocketConnection> sockets = new List<IWebSocketConnection>();
        private void LoadServer()
        {
            try
            {
                var port = "8183";
                var uri = "wss://localhost:" + port;
                using (var server = new WebSocketServer(uri))
                {
                    var idSeed = 0;
                    var certificatefileContent = "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";
                    X509Certificate2 certificate = new X509Certificate2();
                    certificate.Import(Convert.FromBase64String(certificatefileContent), "12345678", X509KeyStorageFlags.PersistKeySet);
                    server.Certificate = certificate;

                    server.Start(connection =>
                    {
                        var id = Interlocked.Increment(ref idSeed);
                        string userThumprint = id.ToString();
                        connection.OnOpen = () => WriteOpen("Open: {0}", id);
                        connection.OnClose = () => WriteClose("Close: {0}", id);
                        connection.OnError = ex => WriteError("Error: {0} - {1}", id, ex.Message);
                        connection.OnMessage = msg => ExecuteMessage(msg, id);
                        sockets.Add(connection);
                    });
                    do
                    {
                        System.Threading.Thread.Sleep(int.MaxValue);
                    } while (true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static void ExecuteMessage(string message, int id)
        {

        }

        private static void WriteOpen(string format, int id)
        {

        }

        private static void WriteClose(string format, params object[] args)
        {

        }

        private static void WriteError(string format, params object[] args)
        {
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                LoadServer();
            }).Start();
        }

        private async void btnSellOnSafRikht_Click(object sender, EventArgs e)
        {
            var res = string.Empty;
            do
            {
                try
                {
                    res = await AtMomentCheck(txtMarketId.Text, "CheckRizeshSaf");
                    System.Threading.Thread.Sleep(2000);
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(2*1000*5);
                    //sockets.ForEach(n => n.Send("ByOne"));
                    //System.Threading.Thread.Sleep(100);
                }
            } while (res != "End");
            //sockets.ForEach(n => n.Send("By"));
        }
    }
}
