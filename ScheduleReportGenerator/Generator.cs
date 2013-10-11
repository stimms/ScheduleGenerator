using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ScheduleReportGenerator
{
    class Generator
    {
        public void Generate(FileInfo fileName)
        {
            using(var package = new ExcelPackage(fileName))
            {
                var worksheet = package.Workbook.Worksheets.Add("Schedule");
                worksheet.Cells[1, 1].Value = "hi simon";

                package.Save();
            }
        }
    }
}
