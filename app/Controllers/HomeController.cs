using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using app.Models;
using System.Net.Http;
using MinMax;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Net;
using OfficeOpenXml.Drawing.Chart;

namespace app.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> IndexOld()
        {
            var ew = await Index1x();
            return View();
        }


        public async Task<IActionResult> FromTseTmc(string Option)
        {
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
                    string exResult = "marketResult.xlsx";
                    FileInfo fileResult = new FileInfo(exResult);
                    using (ExcelPackage result = new ExcelPackage(fileResult))
                    {
                        var mainSheet = marketExcel.Workbook.Worksheets.Where(n => n.Name == "دیده بان بازار").FirstOrDefault();
                        int rowCount = mainSheet.Dimension.Rows;
                        int ColCount = mainSheet.Dimension.Columns;
                        for (int row = 4; row <= rowCount; row++)
                        {
                            var marketName = mainSheet.Cells[row, 1].Value.ToString();
                            if (string.IsNullOrEmpty(marketName) || marketName.Any(c => char.IsDigit(c)) || marketName.EndsWith('ح'))
                                continue;

                            var resultSheet = result.Workbook.Worksheets.Where(n => n.Name == marketName).FirstOrDefault();
                            if (resultSheet == null)
                            {
                                throw new Exception("new market added");
                            }
                            var resultCurrentRowIndex = resultSheet.Dimension.Rows + 1;
                            resultSheet.Cells[resultCurrentRowIndex, 1].Value = this.GetCurrentDate();
                            for (int col = 2; col <= 22; col++)
                            {
                                resultSheet.Cells[resultCurrentRowIndex, col].Value = mainSheet.Cells[row, col + 1].Value;
                            }

                            //var handler = new HttpClientHandler();
                            //handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                            //client = new HttpClient(handler);
                            string searchMarketURL = $"http://tsetmc.com/tsev2/data/search.aspx?skey={marketName}";
                            var findNameResult = await client.GetStringAsync(searchMarketURL);
                            var findedItems = findNameResult.Split(';');
                            var marketId = findedItems.Where(n => n.Split(',')[0] == marketName).Select(n => n.Split(',')[2]).First().ToString();


                            string marketOnlineInfoUrl = $"http://www.tsetmc.com/tsev2/data/instinfodata.aspx?i={marketId}&c=41+";
                            var marketOnlineInfo = await client.GetStringAsync(marketOnlineInfoUrl);
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
                                resultSheet.Cells[resultCurrentRowIndex, 48].Value = double.Parse(hogogiSeal);
                                resultSheet.Cells[resultCurrentRowIndex, 49].Value = double.Parse(hogogiBay);
                                resultSheet.Cells[resultCurrentRowIndex, 50].Value = double.Parse(hagigiBay);
                                resultSheet.Cells[resultCurrentRowIndex, 51].Value = double.Parse(hagigiSeal);
                                resultSheet.Cells[resultCurrentRowIndex, 52].Value = double.Parse(hagmKol);
                                resultSheet.Cells[resultCurrentRowIndex, 53].Value = double.Parse(mablagKol);

                            }
                            else
                            {


                            }

                            //try
                            //{
                            //    marketName = marketName.Replace('ك', 'ک');
                            //    marketName = marketName.Replace('ي', 'ی');


                            //    var stockwatch = $"https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={marketName}&stockWatch=1&";
                            //    //uri = new Uri(stockwatch);
                            //    var sahamyabData = await client.GetStringAsync(stockwatch);
                            //    var sahamyabmarketinfo = Newtonsoft.Json.JsonConvert.DeserializeObject<SahamyabMarketInfo>(sahamyabData);
                            //    resultSheet.Cells[resultCurrentRowIndex, 23].Value = sahamyabmarketinfo.result[0].sahamayb_post_count;
                            //    resultSheet.Cells[resultCurrentRowIndex, 24].Value = sahamyabmarketinfo.result[0].sahamayb_post_count_rank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 25].Value = sahamyabmarketinfo.result[0].sahamyab_follower_count_rank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 26].Value = sahamyabmarketinfo.result[0].sahamyab_page_visit_rank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 27].Value = sahamyabmarketinfo.result[0].marketValueRank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 28].Value = sahamyabmarketinfo.result[0].marketValueRankGroup;
                            //    resultSheet.Cells[resultCurrentRowIndex, 29].Value = sahamyabmarketinfo.result[0].index_affect;
                            //    resultSheet.Cells[resultCurrentRowIndex, 30].Value = sahamyabmarketinfo.result[0].index_affect_rank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 31].Value = sahamyabmarketinfo.result[0].correlation_dollar;
                            //    resultSheet.Cells[resultCurrentRowIndex, 32].Value = sahamyabmarketinfo.result[0].correlation_main_index;
                            //    resultSheet.Cells[resultCurrentRowIndex, 33].Value = sahamyabmarketinfo.result[0].correlation_oil_opec;
                            //    resultSheet.Cells[resultCurrentRowIndex, 34].Value = sahamyabmarketinfo.result[0].correlation_ons_tala;
                            //    resultSheet.Cells[resultCurrentRowIndex, 35].Value = sahamyabmarketinfo.result[0].monthProfitRank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 36].Value = sahamyabmarketinfo.result[0].monthProfitRankGroup;
                            //    resultSheet.Cells[resultCurrentRowIndex, 37].Value = sahamyabmarketinfo.result[0].PE;
                            //    resultSheet.Cells[resultCurrentRowIndex, 38].Value = sahamyabmarketinfo.result[0].sectorPE;
                            //    resultSheet.Cells[resultCurrentRowIndex, 39].Value = sahamyabmarketinfo.result[0].profit7Days;
                            //    resultSheet.Cells[resultCurrentRowIndex, 40].Value = sahamyabmarketinfo.result[0].profit30Days;
                            //    resultSheet.Cells[resultCurrentRowIndex, 41].Value = sahamyabmarketinfo.result[0].profit91Days;
                            //    resultSheet.Cells[resultCurrentRowIndex, 42].Value = sahamyabmarketinfo.result[0].profit182Days;
                            //    resultSheet.Cells[resultCurrentRowIndex, 43].Value = sahamyabmarketinfo.result[0].profit365Days;
                            //    resultSheet.Cells[resultCurrentRowIndex, 44].Value = sahamyabmarketinfo.result[0].profitAllDays;
                            //    resultSheet.Cells[resultCurrentRowIndex, 45].Value = sahamyabmarketinfo.result[0].tradeVolumeRank;
                            //    resultSheet.Cells[resultCurrentRowIndex, 46].Value = sahamyabmarketinfo.result[0].tradeVolumeRankGroup;
                            //    resultSheet.Cells[resultCurrentRowIndex, 47].Value = sahamyabmarketinfo.result[0].zaribNaghdShavandegi;
                            //}
                            //catch (Exception ex)
                            //{
                            //    var error = ex.Message;
                            //    var fullerror = ex.ToString();
                            //}

                        }
                        result.Save();
                    }
                }
                return Json(new { Success = true });
            }
            catch (Exception ex)
            {
                return Json(new { Success = false, Message = ex.ToString() });
            }
        }

        public async Task<IActionResult> FromSahamyab(string Option)
        {
            try
            {
                string exResult = "marketResult.xlsx";
                FileInfo fileResult = new FileInfo(exResult);
                using (ExcelPackage result = new ExcelPackage(fileResult))
                {
                    int resultCurrentRowIndex = 10;
                    foreach (var resultSheet in result.Workbook.Worksheets)
                    {
                        var marketName = resultSheet.Name;
                        if (marketName == "Charts")
                            continue;
                        marketName = marketName.Replace('ك', 'ک');
                        marketName = marketName.Replace('ي', 'ی');

                        var shamyabURL = $"https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={marketName}&stockWatch=1&";
                        var handler = new HttpClientHandler();
                        handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                        var client = new HttpClient(handler);
                        var sahamyabData = await client.GetStringAsync(shamyabURL);
                        var sahamyabmarketinfo = Newtonsoft.Json.JsonConvert.DeserializeObject<SahamyabMarketInfo>(sahamyabData);
                        resultSheet.Cells[resultCurrentRowIndex, 23].Value = sahamyabmarketinfo.result[0].sahamayb_post_count;
                        resultSheet.Cells[resultCurrentRowIndex, 24].Value = sahamyabmarketinfo.result[0].sahamayb_post_count_rank;
                        resultSheet.Cells[resultCurrentRowIndex, 25].Value = sahamyabmarketinfo.result[0].sahamyab_follower_count_rank;
                        resultSheet.Cells[resultCurrentRowIndex, 26].Value = sahamyabmarketinfo.result[0].sahamyab_page_visit_rank;
                        resultSheet.Cells[resultCurrentRowIndex, 27].Value = sahamyabmarketinfo.result[0].marketValueRank;
                        resultSheet.Cells[resultCurrentRowIndex, 28].Value = sahamyabmarketinfo.result[0].marketValueRankGroup;
                        resultSheet.Cells[resultCurrentRowIndex, 29].Value = sahamyabmarketinfo.result[0].index_affect;
                        resultSheet.Cells[resultCurrentRowIndex, 30].Value = sahamyabmarketinfo.result[0].index_affect_rank;
                        resultSheet.Cells[resultCurrentRowIndex, 31].Value = sahamyabmarketinfo.result[0].correlation_dollar;
                        resultSheet.Cells[resultCurrentRowIndex, 32].Value = sahamyabmarketinfo.result[0].correlation_main_index;
                        resultSheet.Cells[resultCurrentRowIndex, 33].Value = sahamyabmarketinfo.result[0].correlation_oil_opec;
                        resultSheet.Cells[resultCurrentRowIndex, 34].Value = sahamyabmarketinfo.result[0].correlation_ons_tala;
                        resultSheet.Cells[resultCurrentRowIndex, 35].Value = sahamyabmarketinfo.result[0].monthProfitRank;
                        resultSheet.Cells[resultCurrentRowIndex, 36].Value = sahamyabmarketinfo.result[0].monthProfitRankGroup;
                        resultSheet.Cells[resultCurrentRowIndex, 37].Value = sahamyabmarketinfo.result[0].PE;
                        resultSheet.Cells[resultCurrentRowIndex, 38].Value = sahamyabmarketinfo.result[0].sectorPE;
                        resultSheet.Cells[resultCurrentRowIndex, 39].Value = sahamyabmarketinfo.result[0].profit7Days;
                        resultSheet.Cells[resultCurrentRowIndex, 40].Value = sahamyabmarketinfo.result[0].profit30Days;
                        resultSheet.Cells[resultCurrentRowIndex, 41].Value = sahamyabmarketinfo.result[0].profit91Days;
                        resultSheet.Cells[resultCurrentRowIndex, 42].Value = sahamyabmarketinfo.result[0].profit182Days;
                        resultSheet.Cells[resultCurrentRowIndex, 43].Value = sahamyabmarketinfo.result[0].profit365Days;
                        resultSheet.Cells[resultCurrentRowIndex, 44].Value = sahamyabmarketinfo.result[0].profitAllDays;
                        resultSheet.Cells[resultCurrentRowIndex, 45].Value = sahamyabmarketinfo.result[0].tradeVolumeRank;
                        resultSheet.Cells[resultCurrentRowIndex, 46].Value = sahamyabmarketinfo.result[0].tradeVolumeRankGroup;
                        resultSheet.Cells[resultCurrentRowIndex, 47].Value = sahamyabmarketinfo.result[0].zaribNaghdShavandegi;
                    }
                    result.Save();

                    return Json(new { Success = true });
                }
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    Success = false,
                    Message = ex.ToString()
                });
            }
        }


        public IActionResult CreateChart()
        {
            try
            {
                this.DoCreateChart();
                return Json(new { Success = true });
            }
            catch (Exception ex)
            {
                return Json(new { Success = false, Message = ex.ToString() });
            }
        }

        public async Task<IActionResult> Index1x()
        {

            //this.CreateChart();
            //return View();
            ////var marketNametest = "فسا";
            //string searchMarketURLb = $"http://www.tsetmc.com/tsev2/data/instinfodata.aspx?i=318005355896147&c=41+";
            ////searchMarketURL = "http://www.tsetmc.com/tsev2/data/instinfofast.aspx?i=46752599569017089&c=31+";
            //var handlerb = new HttpClientHandler();
            //handlerb.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            //var clientb= new HttpClient(handlerb);
            ////var uri = new Uri(searchMarketURL);
            //var search = await clientb.GetStringAsync(searchMarketURLb);
            //var datac = search.Split(';');
            //var a = datac[0];
            //foreach (var itemb in datac)
            //{
            //    var datad = itemb.Split(',');
            //    var d = datad[0];
            //    var e = datad[1];
            //}
            //var c = datac[0];








            string sFileName = @"MarketWatchPlus.xlsx";
            string exResult = "marketResult.xlsx";
            FileInfo file = new FileInfo(sFileName);
            FileInfo fileResult = new FileInfo(exResult);


            using (ExcelPackage marketExcel = new ExcelPackage(file))
            {
                using (ExcelPackage result = new ExcelPackage(fileResult))
                {
                    var mainSheet = marketExcel.Workbook.Worksheets.Where(n => n.Name == "دیده بان بازار").FirstOrDefault();
                    //ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int rowCount = mainSheet.Dimension.Rows;
                    int ColCount = mainSheet.Dimension.Columns;
                    for (int row = 4; row <= rowCount; row++)
                    {
                        var marketName = mainSheet.Cells[row, 1].Value.ToString();
                        if (string.IsNullOrEmpty(marketName) || marketName.Any(c => char.IsDigit(c)) || marketName.EndsWith('ح'))
                            continue;
                        var resultSheet = result.Workbook.Worksheets.Where(n => n.Name == marketName).FirstOrDefault();
                        if (resultSheet == null)
                        {
                            result.Workbook.Worksheets.Add(marketName);
                            resultSheet = result.Workbook.Worksheets.Where(n => n.Name == marketName).FirstOrDefault();
                            resultSheet.Cells[1, 1].Value = "تاریخ";
                            resultSheet.Cells[1, 2].Value = "تعداد";
                            resultSheet.Cells[1, 3].Value = "حجم";
                            resultSheet.Cells[1, 4].Value = "ارزش";
                            resultSheet.Cells[1, 5].Value = "دیروز";
                            resultSheet.Cells[1, 6].Value = "اولین";
                            resultSheet.Cells[1, 7].Value = "آخرین معامله - مقدار";
                            resultSheet.Cells[1, 8].Value = "آخرین معامله - تغییر";
                            resultSheet.Cells[1, 9].Value = "آخرین معامله - درصد";
                            resultSheet.Cells[1, 10].Value = "قیمت پایانی - مقدار";
                            resultSheet.Cells[1, 11].Value = "قیمت پایانی - تغییر";
                            resultSheet.Cells[1, 12].Value = "قیمت پایانی - درصد";
                            resultSheet.Cells[1, 13].Value = "کمترین";
                            resultSheet.Cells[1, 14].Value = "بیشترین";
                            resultSheet.Cells[1, 15].Value = "EPS";
                            resultSheet.Cells[1, 16].Value = "P/E";
                            resultSheet.Cells[1, 17].Value = "خرید - تعداد";
                            resultSheet.Cells[1, 18].Value = "خرید - حجم";
                            resultSheet.Cells[1, 19].Value = "خرید - قیمت";
                            resultSheet.Cells[1, 20].Value = "فروش - قیمت";
                            resultSheet.Cells[1, 21].Value = "فروش - حجم";
                            resultSheet.Cells[1, 22].Value = "فروش - تعداد";
                            //resultSheet.Cells[1, 13].Value = "EPSEPSEPS";
                            //resultSheet.Cells[1, 13].Value = "EPSEPSEPS";
                        }
                        var resultCurrentRowIndex = resultSheet.Dimension.Rows + 1;
                        resultSheet.Cells[resultCurrentRowIndex, 1].Value = this.GetCurrentDate();
                        for (int col = 2; col <= 22; col++)
                        {
                            resultSheet.Cells[resultCurrentRowIndex, col].Value = mainSheet.Cells[row, col + 1].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill= mainSheet.Cells[row, col + 1].Style.Fill;
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill.PatternType = ExcelFillStyle.LightGrid;
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill.BackgroundColor.SetColor(Color.Green);
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill.BackgroundColor.SetColor(this.GetExcelColor( mainSheet.Cells[row, col + 1].Style.Fill.BackgroundColor));
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill.PatternColor.SetColor(this.GetExcelColor(mainSheet.Cells[row, col + 1].Style.Fill.PatternColor));
                            //resultSheet.Cells[resultCurrentRowIndex, col].Style.Fill.PatternColor.SetColor(Color.Green);
                        }
                        //string searchMarketURL = $"http://tsetmc.com/tsev2/data/search.aspx?skey={marketName}";
                        var handler = new HttpClientHandler();
                        handler.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                        var client = new HttpClient(handler);
                        //var findNameResult = await client.GetStringAsync(searchMarketURL);
                        //var findedItems = findNameResult.Split(';');
                        //var marketId=findedItems.Where(n => n.Split(',')[0] == marketName).Select(n => n.Split(',')[2]);

                        //string marketOnlineInfoUrl = $"http://www.tsetmc.com/tsev2/data/instinfodata.aspx?i={marketId}&c=41+";
                        //var marketOnlineInfo = await client.GetStringAsync(marketOnlineInfoUrl);


                        try
                        {
                            marketName = marketName.Replace('ك', 'ک');
                            marketName = marketName.Replace('ي', 'ی');


                            var stockwatch = $"https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={marketName}&stockWatch=1&";
                            //uri = new Uri(stockwatch);
                            var sahamyabData = await client.GetStringAsync(stockwatch);
                            var sahamyabmarketinfo = Newtonsoft.Json.JsonConvert.DeserializeObject<SahamyabMarketInfo>(sahamyabData);
                            resultSheet.Cells[resultCurrentRowIndex, 23].Value = sahamyabmarketinfo.result[0].sahamayb_post_count;
                            resultSheet.Cells[resultCurrentRowIndex, 24].Value = sahamyabmarketinfo.result[0].sahamayb_post_count_rank;
                            resultSheet.Cells[resultCurrentRowIndex, 25].Value = sahamyabmarketinfo.result[0].sahamyab_follower_count_rank;
                            resultSheet.Cells[resultCurrentRowIndex, 26].Value = sahamyabmarketinfo.result[0].sahamyab_page_visit_rank;
                            resultSheet.Cells[resultCurrentRowIndex, 27].Value = sahamyabmarketinfo.result[0].marketValueRank;
                            resultSheet.Cells[resultCurrentRowIndex, 28].Value = sahamyabmarketinfo.result[0].marketValueRankGroup;
                            resultSheet.Cells[resultCurrentRowIndex, 29].Value = sahamyabmarketinfo.result[0].index_affect;
                            resultSheet.Cells[resultCurrentRowIndex, 30].Value = sahamyabmarketinfo.result[0].index_affect_rank;
                            resultSheet.Cells[resultCurrentRowIndex, 31].Value = sahamyabmarketinfo.result[0].correlation_dollar;
                            resultSheet.Cells[resultCurrentRowIndex, 32].Value = sahamyabmarketinfo.result[0].correlation_main_index;
                            resultSheet.Cells[resultCurrentRowIndex, 33].Value = sahamyabmarketinfo.result[0].correlation_oil_opec;
                            resultSheet.Cells[resultCurrentRowIndex, 34].Value = sahamyabmarketinfo.result[0].correlation_ons_tala;
                            resultSheet.Cells[resultCurrentRowIndex, 35].Value = sahamyabmarketinfo.result[0].monthProfitRank;
                            resultSheet.Cells[resultCurrentRowIndex, 36].Value = sahamyabmarketinfo.result[0].monthProfitRankGroup;
                            resultSheet.Cells[resultCurrentRowIndex, 37].Value = sahamyabmarketinfo.result[0].PE;
                            resultSheet.Cells[resultCurrentRowIndex, 38].Value = sahamyabmarketinfo.result[0].sectorPE;
                            resultSheet.Cells[resultCurrentRowIndex, 39].Value = sahamyabmarketinfo.result[0].profit7Days;
                            resultSheet.Cells[resultCurrentRowIndex, 40].Value = sahamyabmarketinfo.result[0].profit30Days;
                            resultSheet.Cells[resultCurrentRowIndex, 41].Value = sahamyabmarketinfo.result[0].profit91Days;
                            resultSheet.Cells[resultCurrentRowIndex, 42].Value = sahamyabmarketinfo.result[0].profit182Days;
                            resultSheet.Cells[resultCurrentRowIndex, 43].Value = sahamyabmarketinfo.result[0].profit365Days;
                            resultSheet.Cells[resultCurrentRowIndex, 44].Value = sahamyabmarketinfo.result[0].profitAllDays;
                            resultSheet.Cells[resultCurrentRowIndex, 45].Value = sahamyabmarketinfo.result[0].tradeVolumeRank;
                            resultSheet.Cells[resultCurrentRowIndex, 46].Value = sahamyabmarketinfo.result[0].tradeVolumeRankGroup;
                            resultSheet.Cells[resultCurrentRowIndex, 47].Value = sahamyabmarketinfo.result[0].zaribNaghdShavandegi;
                        }
                        catch (Exception ex)
                        {
                            var error = ex.Message;
                            var fullerror = ex.ToString();
                        }

                    }
                    result.Save();
                }
            }



            //if (file.Exists)
            //{
            //    file.Delete();
            //    file = new FileInfo(sFileName);
            //}
            //using (ExcelPackage package = new ExcelPackage(file))
            //{
            //    // add a new worksheet to the empty workbook
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employee");
            //    //First add the headers
            //    worksheet.Cells[1, 1].Value = "ID";
            //    worksheet.Cells[1, 2].Value = "Name";
            //    worksheet.Cells[1, 3].Value = "Gender";
            //    worksheet.Cells[1, 4].Value = "Salary (in $)";

            //    //Add values
            //    worksheet.Cells["A2"].Value = 1000;
            //    worksheet.Cells["B2"].Value = "Jon";
            //    worksheet.Cells["C2"].Value = "M";
            //    worksheet.Cells["D2"].Value = 5000;

            //    worksheet.Cells["A3"].Value = 1001;
            //    worksheet.Cells["B3"].Value = "Graham";
            //    worksheet.Cells["C3"].Value = "M";
            //    worksheet.Cells["D3"].Value = 10000;

            //    worksheet.Cells["A4"].Value = 1002;
            //    worksheet.Cells["B4"].Value = "Jenny";
            //    worksheet.Cells["C4"].Value = "F";
            //    worksheet.Cells["D4"].Value = 5000;

            //    package.Save(); //Save the workbook.
            //}







            //string todayDate = "1398/09/20";
            //var todayDirectoryName = todayDate.Replace("/", "_");
            //var url = $"https://www.sahamyab.com/api/proxy/symbol/treeMap?v=0.1&type=volume&market=1,2,4&sector=&timeFrame=day&mini=false&date={todayDate}&";
            //var stockwatch = "https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={0}&stockWatch=1&";
            //var client = new HttpClient();
            //var uri = new Uri(url);
            //string result = await client.GetStringAsync(uri);
            ////string result = System.IO.File.ReadAllText(@"D:\treeMap.json");

            //result = result.Replace("$color", "color");

            //var res = Newtonsoft.Json.JsonConvert.DeserializeObject<SYMain>(result);
            //List<orderdAllData> lstOrdred = new List<orderdAllData>();
            //if (!System.IO.Directory.Exists($"\\{todayDirectoryName}"))
            //    System.IO.Directory.CreateDirectory($"\\{todayDirectoryName}");
            //foreach (var fItem in res.children)
            //{
            //    foreach (var item in fItem.children)
            //    { 
            //        var stockwatchFinalUrl = string.Format(stockwatch, item.name);
            //        string stockwatchResult = await client.GetStringAsync(stockwatchFinalUrl);
            //        System.IO.File.WriteAllText($"\\{todayDirectoryName}\\{item.name}", stockwatchResult);
            //        lstOrdred.Add(new orderdAllData()
            //        {
            //            name = item.name + " \t\t\t " + item.data.description,
            //            percent = item.data.color,
            //        });
            //    }
            //}

            //var final = lstOrdred.OrderBy(n => n.percent);
            //ViewBag.Data = final;

            //System.IO.File.WriteAllText($"\\{todayDirectoryName}\\3m.log", result);
            return View();
        }

        private void DoCreateChart()
        {
            try
            {


                string exResult = "marketResult.xlsx";
                FileInfo fileResult = new FileInfo(exResult);

                using (ExcelPackage result = new ExcelPackage(fileResult))
                {
                    var mainChart = result.Workbook.Worksheets.Where(n => n.Name == "Charts").FirstOrDefault();
                    var excelSampleSheet = result.Workbook.Worksheets.Where(n => n.Name == "فسا").FirstOrDefault();
                    string firstxSeries = excelSampleSheet.Name;

                    ExcelChart visitRanka = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRanka").FirstOrDefault();
                    ExcelChart visitRankb = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRankb").FirstOrDefault();
                    ExcelChart visitRankc = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRankc").FirstOrDefault();
                    if (visitRanka != null)
                    {
                        mainChart.Drawings.Remove(visitRanka);
                        mainChart.Drawings.Remove(visitRankb);
                        mainChart.Drawings.Remove(visitRankc);
                    }
                    visitRanka = mainChart.Drawings.AddChart("SahamyabRanka", eChartType.Line);
                    visitRanka.SetPosition(0, 0, 0, 0);
                    visitRanka.SetSize(500, 400);

                    visitRankb = mainChart.Drawings.AddChart("SahamyabRankb", eChartType.Line);
                    visitRankb.SetPosition(0, 0, 10, 0);
                    visitRankb.SetSize(500, 400);

                    visitRankc = mainChart.Drawings.AddChart("SahamyabRankc", eChartType.Line);
                    visitRankc.SetPosition(0, 0, 20, 0);
                    visitRankc.SetSize(500, 400);

                    var rowCount = excelSampleSheet.Dimension.Rows - 1;
                    var orderedByRankSheets = result.Workbook.Worksheets.Where(n => n.Name != "Charts").OrderBy(n => n.Cells[$"Z{rowCount}"]).ToList();
                    var first20 = orderedByRankSheets.Take(20);

                    first20.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = System.Math.Max(2, max - 30);
                        var series1 = visitRanka.Series.Add($"{n.Name}!Z{min}:Z{max}", $"{firstxSeries}!A{min}:A{max}");
                        series1.Header = n.Name;
                    });

                    ExcelChart visitRank = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabVisitRank").FirstOrDefault();
                    if (visitRank != null)
                        mainChart.Drawings.Remove(visitRank);

                    visitRank = mainChart.Drawings.AddChart("SahamyabVisitRank", eChartType.Line);

                    visitRank.SetPosition(0, 0, 0, 0);
                    visitRank.SetSize(100, 200);



                    var oldChart = mainChart.Drawings.Where(n => n.Name == "chart1").FirstOrDefault();
                    if (oldChart != null)
                        mainChart.Drawings.Remove(oldChart);
                    var diagram = mainChart.Drawings.AddChart("chart1", eChartType.Line);
                    diagram.SetPosition(10, 0, 0, 0);
                    diagram.SetSize(300, 400);

                    foreach (var item in result.Workbook.Worksheets.Where(n => n.Name != "Charts"))
                    {
                        bool canInsert = false;
                        bool canInsertVisitRank = false;
                        int max = item.Dimension.Rows;
                        int min = System.Math.Max(2, max - 30);
                        for (int i = min; i <= max; i++)
                        {
                            int postCount = int.Parse((item.Cells[$"W{i}"].Value ?? "0").ToString());
                            if (postCount > 100)
                            {
                                canInsert = true;
                                //break;
                            }

                            int visitCount = int.Parse((item.Cells[$"Z{i}"].Value ?? "1000").ToString());
                            if (visitCount < 200)
                            {
                                canInsertVisitRank = true;
                                //break;
                            }


                        }
                        if (canInsert)
                        {
                            var series1 = diagram.Series.Add($"{item.Name}!W{min}:W{max}", $"{firstxSeries}!A{min}:A{max}");
                            series1.Header = item.Name;
                        }

                        if (canInsertVisitRank)
                        {
                            var series1 = visitRank.Series.Add($"{item.Name}!Z{min}:Z{max}", $"{firstxSeries}!A{min}:A{max}");
                            series1.Header = item.Name;
                        }
                    }

                    visitRank.Title.Text = "رتبه بازدید";
                    diagram.Title.Text = "تعداد پست سهام یاب";
                    //var diagram = mainChart.Drawings.AddChart("chart1", eChartType.Line);
                    //for (int i = 1; i <= 6; i++)
                    //{
                    //    var series = diagram.Series.Add("فسا!" + $"B{i}:C{i}", "B1:C1");
                    //    //var series = diagram.Series.Add( $"B{i}:C{i}", "B1:C1");
                    //    series.Header = excelWorksheet.Cells[$"A{i}"].Value.ToString();
                    //}
                    diagram.Border.Fill.Color = System.Drawing.Color.Green;
                    result.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("chart", ex);
            }
        }

        //private Color GetExcelColor(ExcelColor backgroundColor)
        //{
        //    System.Drawing.Color CurrentCellColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor.LookupColor());
        //    return CurrentCellColor;
        //}

        private string GetCurrentDate()
        {
            var d = DateTime.Now.AddDays(-1);
            PersianCalendar pc = new PersianCalendar();
            return string.Format("{0}/{1}/{2}", pc.GetYear(d), pc.GetMonth(d).ToString().PadLeft(2, '0'), pc.GetDayOfMonth(d).ToString().PadLeft(2, '0'));
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
