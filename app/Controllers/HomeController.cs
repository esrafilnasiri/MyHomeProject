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
                                //throw new Exception("new market added");
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
                            }
                            var resultCurrentRowIndex = resultSheet.Dimension.Rows + 1;
                            resultSheet.Cells[resultCurrentRowIndex, 1].Value = this.GetCurrentDate();
                            for (int col = 2; col <= 22; col++)
                            {
                                resultSheet.Cells[resultCurrentRowIndex, col].Value = mainSheet.Cells[row, col + 1].Value;
                            }

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
            string exResult = "marketResult.xlsx";
            FileInfo fileResult = new FileInfo(exResult);
            using (ExcelPackage result = new ExcelPackage(fileResult))
            {
                foreach (var resultSheet in result.Workbook.Worksheets)
                {
                    try
                    {
                        var marketName = resultSheet.Name;
                        if (marketName == "Charts")
                            continue;

                        int resultCurrentRowIndex = resultSheet.Dimension.Rows;
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
                    catch (Exception ex)
                    {

                    }
                }
                result.Save();

                return Json(new { Success = true });
            }
        }

        [HttpPost]
        public async Task<IActionResult> FromTseTmcOldDate(string Option)
        {
            try
            {
                string getTsemcExcelURL = "http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=" + Option;
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
                                //throw new Exception("new market added");
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
                            }
                            var resultCurrentRowIndex = resultSheet.Dimension.Rows + 1;
                            resultSheet.Cells[resultCurrentRowIndex, 1].Value = Option;
                            for (int col = 2; col <= 22; col++)
                            {
                                resultSheet.Cells[resultCurrentRowIndex, col].Value = mainSheet.Cells[row, col + 1].Value;
                            }


                            resultSheet.Cells[resultCurrentRowIndex, 23].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 23].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 24].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 24].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 25].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 25].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 26].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 26].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 27].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 27].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 28].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 28].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 29].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 29].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 30].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 30].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 31].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 31].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 32].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 32].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 33].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 33].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 34].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 34].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 35].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 35].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 36].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 36].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 37].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 37].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 38].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 38].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 39].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 39].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 40].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 40].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 41].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 41].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 42].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 42].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 43].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 43].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 44].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 44].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 45].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 45].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 46].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 46].Value;
                            resultSheet.Cells[resultCurrentRowIndex, 47].Value = resultSheet.Cells[resultCurrentRowIndex - 1, 47].Value;


                            //resultSheet.Cells[resultCurrentRowIndex, 23].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 23].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 24].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 24].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 25].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 25].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 26].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 26].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 27].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 27].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 28].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 28].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 29].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 29].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 30].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 30].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 31].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 31].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 32].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 32].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 33].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 33].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 34].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 34].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 35].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 35].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 36].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 36].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 37].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 37].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 38].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 38].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 39].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 39].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 40].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 40].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 41].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 41].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 42].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 42].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 43].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 43].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 44].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 44].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 45].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 45].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 46].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 46].Value;
                            //resultSheet.Cells[resultCurrentRowIndex, 47].Value = resultSheet.Cells[resultCurrentRowIndex - 2, 47].Value;
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
        

        public IActionResult RebuildData()
        {
            try
            {
               
                this.DoRebuildData();
                return Json(new { Success = true});
            }
            catch (Exception ex)
            {
                return Json(new { Success = false, Message = ex.ToString() });
            }
        }

        private void DoRebuildData()
        {
            string exResult = "marketResult.xlsx";
            FileInfo fileResult = new FileInfo(exResult);

            using (ExcelPackage result = new ExcelPackage(fileResult))
            {
                var mainChart = result.Workbook.Worksheets.Where(n => n.Name == "Charts").FirstOrDefault();
                ExcelObjectCompare excelObjectCompare = new ExcelObjectCompare();
                var allMarkentNameConts = result.Workbook.Worksheets.Count;
                foreach (var sahamheet in result.Workbook.Worksheets)
                {
                    int rowCount = sahamheet.Dimension.Rows;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var rankValue = sahamheet.Cells[$"Z{row}"].Value;

                    }
                }

            }
        }

        public IActionResult CreateChart()
        {
            try
            {
                string exResult = "marketResult.xlsx";
                FileInfo fileResult = new FileInfo(exResult);
                var seriesPostCountList = new List<app.Helper.Series>();
                var xAxisPostCount = new app.Helper.XAxis();

                var seriesMaxZarar3DayList = new List<app.Helper.Series>();
                var xAxisMaxZarar3Day = new app.Helper.XAxis();

                var seriesMaxZarar7DayList = new List<app.Helper.Series>();
                var xAxisMaxZarar7Day = new app.Helper.XAxis();

                using (ExcelPackage result = new ExcelPackage(fileResult))
                {
                    var mainChart = result.Workbook.Worksheets.Where(n => n.Name == "Charts").FirstOrDefault();
                    var excelSampleSheet = result.Workbook.Worksheets.Where(n => n.Name == "فسا").FirstOrDefault();
                    string firstxSeries = excelSampleSheet.Name;

                    var sahamyabLastdayRowDiffrence = 0;
                    if (excelSampleSheet.Cells[$"W{excelSampleSheet.Dimension.Rows}"].Value == null)
                        sahamyabLastdayRowDiffrence = 1;
                    ExcelObjectCompare excelObjectCompare = new ExcelObjectCompare();
                    var orderedByPostCountSheets = result.Workbook.Worksheets
                                                             .Where(n => n.Name != "Charts")
                                                             .Where(n => n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"] != null && n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"].Value != null && (double)n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"].Value < 300)
                                                             .OrderBy(n => n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"], excelObjectCompare)
                                                             .ToList();

                    var first20 = orderedByPostCountSheets.Take(20);
                    int seriesOrder = 1;
                    first20.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows - sahamyabLastdayRowDiffrence;
                        int min = System.Math.Max(2, max - 30);
                        this.ReCorrectionOutOfRangeValue(n, min, max, "BB");
                        var data = new List<object>();
                        for (int i = min; i < max; i++)
                        {
                            data.Add(n.Cells[$"BB{i}"].Value);
                            if (seriesOrder == 1)
                                xAxisPostCount.categories.Add(n.Cells[$"A{i}"].Value.ToString());
                        }
                        seriesPostCountList.Add(new app.Helper.Series()
                        {
                            name = n.Name,
                            data = data,
                            visible = seriesOrder < 4,
                        });
                        seriesOrder++;
                    });

                    Double tryCast = 0;
                    result.Workbook.Worksheets
                                   .Where(n => n.Name != "Charts")
                                   .ToList().ForEach(itemSheet =>
                                   {
                                       double mainScore90 = 100;
                                       double mainScore60 = 100;
                                       double mainScore30 = 100;
                                       double mainScore14 = 100;
                                       double mainScore7 = 100;
                                       double mainScore3 = 100;
                                       int sheetRowCount = itemSheet.Dimension.Rows;
                                       int fromRow = Math.Min(90, itemSheet.Dimension.Rows)-2;
                                       for (int i = fromRow; i >= 0; i--)
                                       {
                                           if (itemSheet.Cells[$"L{sheetRowCount - i}"]!=null && itemSheet.Cells[$"L{sheetRowCount - i}"].Value != null && double.TryParse(itemSheet.Cells[$"L{sheetRowCount - i}"].Value.ToString(), out tryCast))
                                           {
                                               var darsad = double.Parse(itemSheet.Cells[$"L{sheetRowCount - i}"].Value.ToString());
                                               if (i < 3)
                                                   mainScore3 += (mainScore3 * darsad) / 100;
                                               if (i < 7)
                                                   mainScore7 += (mainScore7 * darsad) / 100;
                                               if (i < 14)
                                                   mainScore14 += (mainScore14 * darsad) / 100;
                                               if (i < 30)
                                                   mainScore30 += (mainScore30 * darsad) / 100;
                                               if (i < 60)
                                                   mainScore60 += (mainScore60 * darsad) / 100;
                                               if (i < 90)
                                                   mainScore90 += (mainScore90 * darsad) / 100;
                                           }
                                       }
                                       itemSheet.Cells[$"BC{sheetRowCount}"].Value = ((mainScore3 / 100) - 1) * 100;
                                       itemSheet.Cells[$"BD{sheetRowCount}"].Value = ((mainScore7 / 100) - 1) * 100;
                                       itemSheet.Cells[$"BE{sheetRowCount}"].Value = ((mainScore14 / 100) - 1) * 100;
                                       itemSheet.Cells[$"BF{sheetRowCount}"].Value = ((mainScore30 / 100) - 1) * 100;
                                       itemSheet.Cells[$"BG{sheetRowCount}"].Value = ((mainScore60 / 100) - 1) * 100;
                                       itemSheet.Cells[$"BH{sheetRowCount}"].Value = ((mainScore90 / 100) - 1) * 100;
                                   });
                    result.Save();

                    var orderedMaxZarar3Day = result.Workbook.Worksheets
                                                    .Where(n => n.Name != "Charts")
                                                    .Where(n => n.Cells[$"BC{n.Dimension.Rows}"] != null && n.Cells[$"BC{n.Dimension.Rows}"].Value != null && (double)n.Cells[$"BC{n.Dimension.Rows}"].Value < 3000)
                                                    .OrderBy(n => n.Cells[$"BC{n.Dimension.Rows}"], excelObjectCompare)
                                                    .ToList().Take(20);



                    seriesOrder = 1;
                    orderedMaxZarar3Day.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = max - 3;
                        var data = new List<object>();
                        for (int i = min; i <= max; i++)
                        {
                            data.Add(n.Cells[$"BC{i}"].Value);
                            if (seriesOrder == 1)
                                xAxisMaxZarar3Day.categories.Add(n.Cells[$"A{i}"].Value.ToString());
                        }
                        seriesMaxZarar3DayList.Add(new app.Helper.Series()
                        {
                            name = n.Name,
                            data = data,
                            visible = seriesOrder < 4,
                        });
                        seriesOrder++;
                    });


                    var orderedMaxZarar7Day = result.Workbook.Worksheets
                                        .Where(n => n.Name != "Charts")
                                        .Where(n => n.Cells[$"BD{n.Dimension.Rows}"] != null && n.Cells[$"BD{n.Dimension.Rows}"].Value != null && (double)n.Cells[$"BD{n.Dimension.Rows}"].Value < 3000)
                                        .OrderBy(n => n.Cells[$"BD{n.Dimension.Rows}"], excelObjectCompare)
                                        .ToList().Take(20);

                    seriesOrder = 1;
                    orderedMaxZarar7Day.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = max - 7;
                        var data = new List<object>();
                        for (int i = min; i <= max; i++)
                        {
                            data.Add(n.Cells[$"BD{i}"].Value);
                            if (seriesOrder == 1)
                                xAxisMaxZarar7Day.categories.Add(n.Cells[$"A{i}"].Value.ToString());
                        }
                        seriesMaxZarar7DayList.Add(new app.Helper.Series()
                        {
                            name = n.Name,
                            data = data,
                            visible = seriesOrder < 4,
                        });
                        seriesOrder++;
                    });
                }

                var newChart = new app.Helper.HighChart()
                {
                    title = new Helper.HTitle() { text = "تعداد پست" },
                    subtitle = new Helper.HTitle() { text = "" },
                    yAxis = new Helper.YAxis() { title = new Helper.HTitle() { text = "تعداد" } },
                    legend = new Helper.Legend() { align = "right", layout = "vertical", verticalAlign = "middle" },
                    plotOptions = new Helper.PlotOptions() { series = new Helper.PlotOptionsSeries() { label = new Helper.PlotOptionsSeriesLabel() { connectorAllowed = false } } },
                    series = seriesPostCountList,
                    xAxis = xAxisPostCount
                };

                var maxZarar3Day = new app.Helper.HighChart()
                {
                    title = new Helper.HTitle() { text = "بیشترین ضرر ۳ روزه" },
                    subtitle = new Helper.HTitle() { text = "" },
                    yAxis = new Helper.YAxis() { title = new Helper.HTitle() { text = "ضرر تجمعی" } },
                    legend = new Helper.Legend() { align = "right", layout = "vertical", verticalAlign = "middle" },
                    plotOptions = new Helper.PlotOptions() { series = new Helper.PlotOptionsSeries() { label = new Helper.PlotOptionsSeriesLabel() { connectorAllowed = false } } },
                    series = seriesMaxZarar3DayList,
                    xAxis = xAxisMaxZarar3Day
                };


                var maxZarar7Day = new app.Helper.HighChart()
                {
                    title = new Helper.HTitle() { text = "بیشترین ضرر ۷ روزه" },
                    subtitle = new Helper.HTitle() { text = "" },
                    yAxis = new Helper.YAxis() { title = new Helper.HTitle() { text = "ضرر تجمعی" } },
                    legend = new Helper.Legend() { align = "right", layout = "vertical", verticalAlign = "middle" },
                    plotOptions = new Helper.PlotOptions() { series = new Helper.PlotOptionsSeries() { label = new Helper.PlotOptionsSeriesLabel() { connectorAllowed = false } } },
                    series = seriesMaxZarar7DayList,
                    xAxis = xAxisMaxZarar7Day
                };

                //this.DoCreateChart();
                return Json(new { Success = true, ChartData = newChart , maxZarar3Day, maxZarar7Day });
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
                        resultSheet.Cells[resultCurrentRowIndex, 1].Value = "1398/10/15";//this.GetCurrentDate();
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

                                                                                                                       
                    ExcelChart postCounta = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRanka").FirstOrDefault();
                    ExcelChart visitRankb = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRankb").FirstOrDefault();
                    ExcelChart visitRankc = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabRankc").FirstOrDefault();
                    if (postCounta != null)
                    {
                        mainChart.Drawings.Remove(postCounta);
                        mainChart.Drawings.Remove(visitRankb);
                        mainChart.Drawings.Remove(visitRankc);
                    }
                    //postCounta = mainChart.Drawings.AddChart("SahamyabRanka", eChartType.Line3D);
                    postCounta = mainChart.Drawings.AddChart("SahamyabRanka", eChartType.Line3D);
                    postCounta.SetPosition(0, 0, 0, 0);
                    postCounta.SetSize(500, 400);

                    visitRankb = mainChart.Drawings.AddChart("SahamyabRankb", eChartType.Line3D);
                    visitRankb.SetPosition(10, 0, 10, 0);
                    visitRankb.SetSize(500, 400);

                    visitRankc = mainChart.Drawings.AddChart("SahamyabRankc", eChartType.Line3D);
                    visitRankc.SetPosition(20, 0, 20, 0);
                    visitRankc.SetSize(500, 400);

                    var sahamyabLastdayRowDiffrence = 0;
                    if (excelSampleSheet.Cells[$"W{excelSampleSheet.Dimension.Rows}"].Value == null)
                        sahamyabLastdayRowDiffrence = 1;
                    ExcelObjectCompare excelObjectCompare = new ExcelObjectCompare();
                    var orderedByPostCountSheets = result.Workbook.Worksheets
                                                             .Where(n => n.Name != "Charts")
                                                             .Where(n => n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"] != null && n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"].Value != null && (double)n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"].Value < 300)
                                                             .OrderBy(n => n.Cells[$"X{n.Dimension.Rows - sahamyabLastdayRowDiffrence}"], excelObjectCompare)
                                                             .ToList();

                    var first20 = orderedByPostCountSheets.Take(20);

                    first20.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows-sahamyabLastdayRowDiffrence;
                        int min = System.Math.Max(2, max - 30);
                        this.ReCorrectionOutOfRangeValue(n, min, max, "BB");
                        var series1 = postCounta.Series.Add($"{n.Name}!BB{min}:BB{max}", $"{firstxSeries}!A{min}:A{max}");
                        series1.Header = n.Name;
                    });
                    postCounta.Title.Text = "رتبه تعداد پست در سهام یاب_1";
                    postCounta.YAxis.MaxValue = 300;
                    postCounta.YAxis.MinValue = 200;

                    var second20 = orderedByPostCountSheets.Skip(20).Take(20);
                    second20.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = System.Math.Max(2, max - 30);
                        this.ReCorrectionOutOfRangeValue(n, min, max, "BB");
                        var series1 = visitRankb.Series.Add($"{n.Name}!BB{min}:BB{max}", $"{firstxSeries}!A{min}:A{max}");
                        series1.Header = n.Name;
                    });
                    visitRankb.Title.Text = "رتبه تعداد پست در سهام یاب_2";


                    var three20 = orderedByPostCountSheets.Skip(40).Take(20);
                    three20.ToList().ForEach(n =>
                    {
                        int max = n.Dimension.Rows;
                        int min = System.Math.Max(2, max - 30);
                        this.ReCorrectionOutOfRangeValue(n, min, max, "BB");
                        var series1 = visitRankc.Series.Add($"{n.Name}!BB{min}:BB{max}", $"{firstxSeries}!A{min}:A{max}");
                        series1.Header = n.Name;
                    });
                    visitRankc.Title.Text = "رتبه تعداد پست در سهام یاب_3";




                    //gabl manfi akharing balaye 4.8
                    double tryCast = 0;
                    var akharingMoameleBalaye48s = result.Workbook.Worksheets
                                            .Where(n => n.Name != "Charts")
                                            .Where(n=> n.Cells[$"A{n.Dimension.Rows}"]?.Value?.ToString() == this.GetCurrentDate())
                                            .Where(n => n.Cells[$"I{n.Dimension.Rows}"] != null && n.Cells[$"I{n.Dimension.Rows}"].Value != null && double.TryParse(n.Cells[$"I{n.Dimension.Rows}"].Value.ToString(),out tryCast) && double.Parse(n.Cells[$"I{n.Dimension.Rows}"].Value.ToString()) > 4.8 && n.Cells[$"L{n.Dimension.Rows- 1}"] != null && n.Cells[$"L{n.Dimension.Rows - 1}"].Value != null && double.TryParse(n.Cells[$"L{n.Dimension.Rows - 1}"].Value.ToString(), out tryCast) && double.Parse(n.Cells[$"L{n.Dimension.Rows - 1}"].Value.ToString()) < 0)
                                            .OrderBy(n => n.Cells[$"I{n.Dimension.Rows}"], excelObjectCompare).ToList().Take(20);

                    int rowNum = 32;
                    mainChart.Cells[$"A{rowNum}"].Value = "نام نماد";
                    mainChart.Cells[$"B{rowNum}"].Value = "آخرین معامله";
                    mainChart.Cells[$"C{rowNum}"].Value = "پایانی درصد";
                    mainChart.Cells[$"D{rowNum}"].Value = "روز قبل پایانی درصد";
                    mainChart.Cells[$"E{rowNum}"].Value = "سود و زیان ۷ روزه";
                    mainChart.Cells[$"F{rowNum++}"].Value = "سود و زیان ۱۴ روزه";
                    for (int i = 1; i < 20; i++)
                    {
                        mainChart.Cells[$"A{rowNum + i}"].Value = "";
                        mainChart.Cells[$"B{rowNum + i}"].Value = "";
                        mainChart.Cells[$"C{rowNum + i}"].Value = "";
                        mainChart.Cells[$"D{rowNum + i}"].Value = "";
                        mainChart.Cells[$"E{rowNum + i}"].Value = "";
                        mainChart.Cells[$"F{rowNum + i}"].Value = "";
                    }
                    foreach (var item in akharingMoameleBalaye48s)
                    {
                        mainChart.Cells[$"A{rowNum}"].Value = item.Name;
                        mainChart.Cells[$"B{rowNum}"].Value = item.Cells[$"I{item.Dimension.Rows }"].Value;
                        mainChart.Cells[$"C{rowNum}"].Value = item.Cells[$"L{item.Dimension.Rows }"].Value;
                        mainChart.Cells[$"D{rowNum}"].Value = item.Cells[$"L{item.Dimension.Rows - 1}"].Value;

                        //miangin 7 roze
                        double mainScore = 100;
                        var days = item.Dimension.Rows - 1;
                        for (int i = 7; i > 0; i--)
                        {
                            if (item.Cells[$"L{item.Dimension.Rows - i}"].Value != null && double.TryParse(item.Cells[$"L{item.Dimension.Rows - i}"].Value.ToString(), out tryCast))
                            {
                                var darsad = double.Parse(item.Cells[$"L{item.Dimension.Rows - i}"].Value.ToString());
                                mainScore = mainScore + (mainScore * darsad) / 100;
                            }
                        }
                        mainChart.Cells[$"E{rowNum}"].Value = ((mainScore / 100) - 1) * 100;

                        mainScore = 100;
                        days = item.Dimension.Rows - 1;
                        for (int i = 14; i > 0; i--)
                        {
                            if (item.Cells[$"L{item.Dimension.Rows - i}"].Value != null && double.TryParse(item.Cells[$"L{item.Dimension.Rows - i}"].Value.ToString(), out tryCast))
                            {
                                var darsad = double.Parse(item.Cells[$"L{item.Dimension.Rows - i}"].Value.ToString());
                                mainScore = mainScore + (mainScore * darsad) / 100;
                            }
                        }
                        mainChart.Cells[$"F{rowNum}"].Value = ((mainScore / 100) - 1) * 100;


                        rowNum++;
                    }

                    //ExcelChart visitRank = (ExcelChart)mainChart.Drawings.Where(n => n.Name == "SahamyabVisitRank").FirstOrDefault();
                    //if (visitRank != null)
                    //    mainChart.Drawings.Remove(visitRank);

                    //visitRank = mainChart.Drawings.AddChart("SahamyabVisitRank", eChartType.Line);

                    //visitRank.SetPosition(0, 0, 0, 0);
                    //visitRank.SetSize(100, 200);



                    //var oldChart = mainChart.Drawings.Where(n => n.Name == "chart1").FirstOrDefault();
                    //if (oldChart != null)
                    //    mainChart.Drawings.Remove(oldChart);
                    //var diagram = mainChart.Drawings.AddChart("chart1", eChartType.Line);
                    //diagram.SetPosition(10, 0, 0, 0);
                    //diagram.SetSize(300, 400);

                    //foreach (var item in result.Workbook.Worksheets.Where(n => n.Name != "Charts"))
                    //{
                    //    bool canInsert = false;
                    //    bool canInsertVisitRank = false;
                    //    int max = item.Dimension.Rows;
                    //    int min = System.Math.Max(2, max - 30);
                    //    for (int i = min; i <= max; i++)
                    //    {
                    //        int postCount = int.Parse((item.Cells[$"W{i}"].Value ?? "0").ToString());
                    //        if (postCount > 100)
                    //        {
                    //            canInsert = true;
                    //            //break;
                    //        }

                    //        int visitCount = int.Parse((item.Cells[$"Z{i}"].Value ?? "1000").ToString());
                    //        if (visitCount < 200)
                    //        {
                    //            canInsertVisitRank = true;
                    //            //break;
                    //        }


                    //    }
                    //    if (canInsert)
                    //    {
                    //        var series1 = diagram.Series.Add($"{item.Name}!W{min}:W{max}", $"{firstxSeries}!A{min}:A{max}");
                    //        series1.Header = item.Name;
                    //    }

                    //    if (canInsertVisitRank)
                    //    {
                    //        var series1 = visitRank.Series.Add($"{item.Name}!Z{min}:Z{max}", $"{firstxSeries}!A{min}:A{max}");
                    //        series1.Header = item.Name;
                    //    }
                    //}

                    //visitRank.Title.Text = "رتبه بازدید";
                    //diagram.Title.Text = "تعداد پست سهام یاب";
                    ////var diagram = mainChart.Drawings.AddChart("chart1", eChartType.Line);
                    ////for (int i = 1; i <= 6; i++)
                    ////{
                    ////    var series = diagram.Series.Add("فسا!" + $"B{i}:C{i}", "B1:C1");
                    ////    //var series = diagram.Series.Add( $"B{i}:C{i}", "B1:C1");
                    ////    series.Header = excelWorksheet.Cells[$"A{i}"].Value.ToString();
                    ////}
                    //diagram.Border.Fill.Color = System.Drawing.Color.Green;
                    result.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("chart", ex);
            }
        }

        private void ReCorrectionOutOfRangeValue(ExcelWorksheet sheet, int min, int max, string copyResultColumnName)
        {
            float sum = 0;
            int includeSumCount = 0;
            for (int i = min; i <= max; i++)
            {
                var data = sheet.Cells[$"X{i}"].Value;
                if (data != null)
                {
                    sum += float.Parse(data.ToString());
                    includeSumCount++;
                }
            }
            var avr = sum / includeSumCount;

            for (int i = min; i <= max; i++)
            {
                var val = sheet.Cells[$"X{i}"].Value;
                if (val == null || float.Parse(val.ToString()) + 150 < avr || float.Parse(val.ToString()) - 150 > avr)
                    sheet.Cells[$"{copyResultColumnName}{i}"].Value = 300 - avr;
                else
                    sheet.Cells[$"{copyResultColumnName}{i}"].Value = 300 - float.Parse( val.ToString());

            }
        }

        //private Color GetExcelColor(ExcelColor backgroundColor)
        //{
        //    System.Drawing.Color CurrentCellColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor.LookupColor());
        //    return CurrentCellColor;
        //}

        private string GetCurrentDate()
        {
            var d = DateTime.Now.AddDays(0);
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
