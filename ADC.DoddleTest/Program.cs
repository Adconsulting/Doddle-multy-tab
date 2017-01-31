using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using ADC.DoddleTest.Infrastructure.Extensions;
using DoddleReport;
using DoddleReport.OpenXml;

namespace ADC.DoddleTest
{
    class Program
    {
        static void Main(string[] args)
        {

            //Create multitab excel
            DoddleMultiTab();
            //Create excel based on a List<dynamic>
            //DoddleDynamic();
            //Create excel based on a list<array>
            //DoddleStringArrayList();

        }

        private static void BuildExcel()
        {
            var list = GenerateLargeList(50);

            var writer = new ExcelReportWriter();
            var report = list.Export()
                .Column("Naam 1", x => x.Param1)
                .Column("naam 2", x => x.Param2).ToReport(null, new ExcelReportWriter());

            
            using (var sr = new StreamWriter(@"c:\temp\builder-report.xlsx"))
            {
                writer.WriteReport(report, sr.BaseStream);
            }
        }

        private static void DoddleDynamic()
        {
            var list = GetDynamicList();

            var mainreport = new Report(list.ToReportSource(), new ExcelReportWriter());
            mainreport.RenderHints["SheetName"] = "dynamic";


            mainreport.DataFields["Col1"].HeaderText = "Col - 1";
            var writer = new ExcelReportWriter();

            using (var sr = new StreamWriter(@"c:\temp\dynamic-report.xlsx"))
            {
                writer.WriteReport(mainreport, sr.BaseStream);
            }
        }

        private static void DoddleStringArrayList()
        {
            string[] header;
            var list = GetArrayList(10, out header);

            var headers = new List<KeyValuePair<string, string>>();
            headers.Add(new KeyValuePair<string, string>("Datum", new DateTime().ToString("dd/MM/yyyy")));

            var export = list.Export();

            var cols = list.First();

            for (int i = 0; i < header.Length ; i++)
            {
                var i1 = i;
                export = export.Column($"{header[i1]}", x => x[i1]);
            }




            //.Column("col 1", x => x[0])
            //.Column("Col 2", x => x[1]).ToReport(headers);
            var mainreport = export.ToReport(null);
            mainreport.RenderHints["SheetName"] = "string-array";
            var writer = new ExcelReportWriter();

            using (var sr = new StreamWriter(@"c:\temp\string-array-report.xlsx"))
            {
                writer.WriteReport(mainreport, sr.BaseStream);
            }
        }

        private static void DoddleMultiTab()
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
                writer.WriteReport(mainreport, sr.BaseStream);
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

        private static List<ExpandoObject> GetDynamicList()
        {
            var retVal = new List<ExpandoObject>();

            dynamic val1 = new ExpandoObject();
            val1.Col1 = "col1 content";
            val1.Col2 = "col2 content";

            dynamic val2 = new ExpandoObject();
            val2.Col1 = "col1 content";
            val2.Col2 = "col2 content";

            dynamic val3 = new ExpandoObject();
            val3.Col1 = "col1 content";
            val3.Col2 = "col2 content";

            dynamic val4 = new ExpandoObject();
            val4.Col1 = "col1 content";
            val4.Col2 = "col2 content";
            val4.Col3 = "col3 content";

            retVal.Add(val4);
            retVal.Add(val2);
            retVal.Add(val3);
            retVal.Add(val1);


            return retVal;
        }

        private static List<string[]> GetArrayList(int nrOfItems, out string[] header)
        {
            var list = new List<string[]>();

            for (int i = 0; i < nrOfItems; i++)
            {
                list.Add(new string[] { $"row {i}", "string 1", "string 2", "string 3", "string 4" });
            }

            header = new List<string> {"Rij", "tekst 1", "tekst 2", "tekst 3", "tekst 4"}.ToArray();

            return list;
        }
    }
}
