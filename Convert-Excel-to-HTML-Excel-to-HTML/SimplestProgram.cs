using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Convert_Excel_to_HTML_Excel_to_HTML
{
    class SimplestProgram
    {
        static void Main(string[] args)
        {
            try
            {
                //load Excel file
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(@"test.xlsx");

                //convert Excel to HTML
                Worksheet sheet = workbook.Worksheets[0];
                sheet.SaveToHtml("sample.html");

                //Preview HTML
                System.Diagnostics.Process.Start("sample.html");
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
