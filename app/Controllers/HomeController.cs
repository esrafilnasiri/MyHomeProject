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

namespace app.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            string sFileName = @"market.xlsx";
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
                        if (string.IsNullOrEmpty(marketName) || marketName.Any(c => char.IsDigit(c)))
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
                        string searchMarketURL = $"http://tsetmc.com/tsev2/data/search.aspx?skey={marketName}";
                        searchMarketURL = "http://www.tsetmc.com/tsev2/data/instinfofast.aspx?i=46752599569017089&c=31+";
                        var client = new HttpClient();
                        var uri = new Uri(searchMarketURL);
                        var res = await client.GetAsync(uri);
                        var vv = res.Content.ReadAsStringAsync().Result;
                        //string searchResult = Encoding.UTF8.GetString(res);
                        //var splitResult = searchResult.Split(',');

                        try
                        {

                            var stockwatch = $"https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={marketName}&stockWatch=1&";
                            uri = new Uri(stockwatch);
                            var sahamyabData = await client.GetStringAsync(uri);
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

        private Color GetExcelColor(ExcelColor backgroundColor)
        {
            System.Drawing.Color CurrentCellColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor.LookupColor());
            return CurrentCellColor;
        }

        private string GetCurrentDate()
        {
            var d = DateTime.Now;
            PersianCalendar pc = new PersianCalendar();
            return string.Format("{0}/{1}/{2}", pc.GetYear(d), pc.GetMonth(d), pc.GetDayOfMonth(d));
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
