using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DoddleReport;
using DoddleReport.OpenXml;

namespace ADC.DoddleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var list1 = GenerateList(4);
            var list2 = GenerateLargeList(150);
            var list3 = GenerateList(10);


            var mainreport = new Report(list1.ToReportSource(), new ExcelReportWriter());
            mainreport.RenderHints["SheetName"] = "List 1";

            var report2 = new Report(list2.ToReportSource(), new ExcelReportWriter());
            report2.RenderHints["SheetName"] = "List 2";

            var report3 = new Report(list3.ToReportSource(), new ExcelReportWriter());
            report3.RenderHints["SheetName"] = "List 3";

            var writer = new ExcelReportWriter();
            writer.AppendReport(mainreport, report2);
            writer.AppendReport(mainreport, report3);

            using (var sr = new StreamWriter(@"c:\temp\report.xlsx"))
            {
                writer.WriteReport(mainreport,sr.BaseStream);
            }

            


        }

        private static List<AdcObject> GenerateList(int nrOfItems)
        {
            var list = new List<AdcObject>();

            for (int i = 0; i <= nrOfItems; i++)
            {
                list.Add(new AdcObject
                {
                    Param1 = string.Format(@"{0}-0", i),
                    Param2 = string.Format(@"{0}-1", i),
                    Param3 = string.Format(@"{0}-2", i),
                    Param4 = i
                });
            }
            return list;

        }

        private static List<AdcLargerObject> GenerateLargeList(int nrOfItems)
        {
            var list = new List<AdcLargerObject>();

            for (int i = 0; i <= nrOfItems; i++)
            {
                list.Add(new AdcLargerObject
                {
                    Param1 = string.Format(@"{0}-0", i),
                    Param2 = string.Format(@"{0}-1", i),
                    Param3 = string.Format(@"{0}-2", i),
                    Param4 = i,
                    Param5 = string.Format(@"{0}-3", i),
                    Param6 = string.Format(@"{0}-4", i),
                    Param7 = string.Format(@"{0}-5", i),
                    Param8 = i
                });
            }
            return list;

        } 
    }
}
