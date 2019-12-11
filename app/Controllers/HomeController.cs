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
            string todayDate = "1398/09/20";
            var todayDirectoryName = todayDate.Replace("/", "_");
            var url = $"https://www.sahamyab.com/api/proxy/symbol/treeMap?v=0.1&type=volume&market=1,2,4&sector=&timeFrame=day&mini=false&date={todayDate}&";
            var stockwatch = "https://www.sahamyab.com/api/proxy/symbol/getSymbolExtData?v=0.1&code={0}&stockWatch=1&";
            var client = new HttpClient();
            var uri = new Uri(url);
            string result = await client.GetStringAsync(uri);
            //string result = System.IO.File.ReadAllText(@"D:\treeMap.json");

            result = result.Replace("$color", "color");

            var res = Newtonsoft.Json.JsonConvert.DeserializeObject<SYMain>(result);
            List<orderdAllData> lstOrdred = new List<orderdAllData>();
            if (!System.IO.Directory.Exists($"\\{todayDirectoryName}"))
                System.IO.Directory.CreateDirectory($"\\{todayDirectoryName}");
            foreach (var fItem in res.children)
            {
                foreach (var item in fItem.children)
                { 
                    var stockwatchFinalUrl = string.Format(stockwatch, item.name);
                    string stockwatchResult = await client.GetStringAsync(stockwatchFinalUrl);
                    System.IO.File.WriteAllText($"\\{todayDirectoryName}\\{item.name}", stockwatchResult);
                    lstOrdred.Add(new orderdAllData()
                    {
                        name = item.name + " \t\t\t " + item.data.description,
                        percent = item.data.color,
                    });
                }
            }

            var final = lstOrdred.OrderBy(n => n.percent);
            ViewBag.Data = final;

            System.IO.File.WriteAllText($"\\{todayDirectoryName}\\3m.log", result);
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
