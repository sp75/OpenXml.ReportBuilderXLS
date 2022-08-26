using DocumentFormat.OpenXml.ReportBuilder;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var list = new[] { new { Code = 1 } }.ToList();

            for (int i = 2; i < 10; i++)
            {
                list.Add(new { Code = i });
            }

            var ddd = ReportBuilderXLS.GenerateReport(new Dictionary<string, IList> { { "www", list } }, @"c:\WinVSProjects\OpenXml.ReportBuilderXLS\TestReport.xlsx").ToArray();




            File.WriteAllBytes(@"c:\Temp\TestReport_rezult.xlsx", ddd);

            if (File.Exists(@"c:\Temp\TestReport_rezult.xlsx"))
            {
                Process.Start(@"c:\Temp\TestReport_rezult.xlsx");
            }


        }
    }
}
