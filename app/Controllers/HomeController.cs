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
            string sFileName = @"demo.xlsx";
            FileInfo file = new FileInfo(sFileName);
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(sFileName));
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employee");
                //First add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Gender";
                worksheet.Cells[1, 4].Value = "Salary (in $)";

                //Add values
                worksheet.Cells["A2"].Value = 1000;
                worksheet.Cells["B2"].Value = "Jon";
                worksheet.Cells["C2"].Value = "M";
                worksheet.Cells["D2"].Value = 5000;

                worksheet.Cells["A3"].Value = 1001;
                worksheet.Cells["B3"].Value = "Graham";
                worksheet.Cells["C3"].Value = "M";
                worksheet.Cells["D3"].Value = 10000;

                worksheet.Cells["A4"].Value = 1002;
                worksheet.Cells["B4"].Value = "Jenny";
                worksheet.Cells["C4"].Value = "F";
                worksheet.Cells["D4"].Value = 5000;

                package.Save(); //


            }

                //var newFile = @"newbook.core.xlsx";

                //using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                //{

                //    IWorkbook workbook = new XSSFWorkbook();

                //    ISheet sheet1 = workbook.CreateSheet("Sheet1");

                //    sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
                //    var rowIndex = 0;
                //    IRow row = sheet1.CreateRow(rowIndex);
                //    row.Height = 30 * 80;
                //    row.CreateCell(0).SetCellValue("this is content");
                //    sheet1.AutoSizeColumn(0);
                //    rowIndex++;

                //    var sheet2 = workbook.CreateSheet("Sheet2");
                //    var style1 = workbook.CreateCellStyle();
                //    style1.FillForegroundColor = HSSFColor.Blue.Index2;
                //    style1.FillPattern = FillPattern.SolidForeground;

                //    var style2 = workbook.CreateCellStyle();
                //    style2.FillForegroundColor = HSSFColor.Yellow.Index2;
                //    style2.FillPattern = FillPattern.SolidForeground;

                //    var cell2 = sheet2.CreateRow(0).CreateCell(0);
                //    cell2.CellStyle = style1;
                //    cell2.SetCellValue(0);

                //    cell2 = sheet2.CreateRow(1).CreateCell(0);
                //    cell2.CellStyle = style2;
                //    cell2.SetCellValue(1);

                //    workbook.Write(fs);
                //}

                //var tsemcePath = "http://members.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0";
                //var client = new HttpClient();
                //var uri = new Uri(tsemcePath);
                //var excelResult = await client.GetStreamAsync(uri);

                //IWorkbook workbook = new XSSFWorkbook(@"C:\Users\Nasiri\Downloads\MarketWatchPlus-1398_9_23.xlsx");
                //using (FileStream stream = new FileStream(@"D:\Book2.xlsx", FileMode.Create, FileAccess.ReadWrite))
                //{
                //    IWorkbook workbookResult = new XSSFWorkbook();
                //    var index = workbook.GetSheetIndex("دیده بان بازار");
                //    ISheet sheet = workbook.GetSheetAt(index);
                //    for (int row = 3; row <= sheet.LastRowNum; row++)
                //    {
                //        var theRow = sheet.GetRow(row);
                //        string namadName = theRow.GetCell(0).StringCellValue;
                //        if (string.IsNullOrEmpty(namadName))
                //            continue;
                //        Regex r = new Regex(@"\d+");
                //        if (r.IsMatch(namadName))
                //            continue;

                //        int indexResult = workbookResult.GetSheetIndex(namadName);
                //        ISheet sheet1;
                //        if (indexResult == -1)
                //            sheet1 = workbookResult.CreateSheet(namadName);
                //        else
                //            sheet1 = workbookResult.GetSheetAt(index);
                //        IRow newrow = sheet1.CreateRow(sheet1.LastRowNum);
                //        //var cell = newrow.CreateCell(0);
                //        //cell.SetCellValue(theRow.GetCell(1).StringCellValue);
                //        break;
                //        //if (sheet.GetRow(row) != null)
                //        //{
                //        //    MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                //        //}
                //    }

                //    using (var file2 = new FileStream("D:\\banding2.xlsx", FileMode.Create, FileAccess.ReadWrite))
                //    {
                //        workbookResult.Write(file2);
                //        file2.Close();
                //    }
                //    workbookResult.Write(stream);
                //}

                //using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                //{





                //IWorkbook workbook = new XSSFWorkbook();
                //workbook.GetSheetIndex("");
                //ISheet sheet1 = workbook.CreateSheet("Sheet1");

                ////sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
                //var rowIndex = 0;
                //IRow row = sheet1.CreateRow(rowIndex);
                //row.Height = 30 * 80;
                //row.CreateCell(0).SetCellValue("this is content");
                //sheet1.AutoSizeColumn(0);
                //rowIndex++;

                //var sheet2 = workbook.CreateSheet("Sheet2");
                //var style1 = workbook.CreateCellStyle();
                ////style1.FillForegroundColor = HSSFColor.Blue.Index2;
                //style1.FillPattern = FillPattern.SolidForeground;

                //var style2 = workbook.CreateCellStyle();
                ////style2.FillForegroundColor = HSSFColor.Yellow.Index2;
                //style2.FillPattern = FillPattern.SolidForeground;

                //var cell2 = sheet2.CreateRow(0).CreateCell(0);
                //cell2.CellStyle = style1;
                //cell2.SetCellValue(0);

                //cell2 = sheet2.CreateRow(1).CreateCell(0);
                //cell2.CellStyle = style2;
                //cell2.SetCellValue(1);

                //workbook.Write(fs);
                //}



                //string todayDate = "1398/09/23";
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
