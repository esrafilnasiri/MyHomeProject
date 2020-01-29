using System;
using System.Collections.Generic;

namespace app.Helper
{
    public class HighChart
    {
        public HighChart()
        {
        }

        public HTitle title { get; set; }
        public HTitle subtitle { get; set; }
        public YAxis yAxis { get; set; }
        public XAxis xAxis { get; set; }
        public Legend legend { get; set; }
        public PlotOptions plotOptions { get; set; }
        public List<Series> series { get; set; }
    }

    public class HTitle
    {
        public string text { get; set; }
    }

    public class YAxis
    {
        public HTitle title { get; set; }
    }

    public class XAxis
    {
        public XAxis()
        {
            categories = new List<string>();
        }
        public List<string> categories { get; set; }
    }

    public class Legend
    {
        public string layout { get; set; }
        public string align { get; set; }
        public string verticalAlign { get; set; }
    }

    public class PlotOptions
    {
        public PlotOptionsSeries series { get; set; }
    }

    public class PlotOptionsSeries
    {
        public  PlotOptionsSeriesLabel label { get;set;}
        public int pointStart { get; set; }
    }

    public class PlotOptionsSeriesLabel
    {
        public bool connectorAllowed { get; set; }
    }

    public class Series
    {
        public string name { get; set; }
        public List<object> data { get; set; }
        public bool visible { get; set; }
    }
}
