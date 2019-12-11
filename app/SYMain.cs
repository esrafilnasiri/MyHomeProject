using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MinMax
{

    public class orderdAllData
    {
        public string name { get; set; }
        public double percent { get; set; }
    }

    public class SYMain
    {
        public string id { get; set; }
        public string name { get; set; }
        public List<Children> children { get; set; }
    }

    public class Children
    {
        public string id { get; set; }
        public string name { get; set; }
        public List<Children> children { get; set; }
        public Data data { get; set; }
    }

    public class Data
    {
        public string description { get; set; }
        public double price { get; set; }
        public double color { get; set; }
    }
}
