using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace app
{
    public class ExcelObjectCompare : IComparer<ExcelRange>
    {
        public int Compare(ExcelRange excelRangeA, ExcelRange excelRangeB)
        {
            Double valueA = double.Parse((excelRangeA.Value ?? "0").ToString());
            Double valueB = double.Parse((excelRangeB.Value ?? "0").ToString());

            if (valueA > valueB)
            {
                return 1;
            }
            else if (valueA < valueB)
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }
    }
}
