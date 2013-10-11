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
            using (var package = new ExcelPackage(fileName))
            {
                var worksheet = package.Workbook.Worksheets.Add("Schedule");
                AddTitles(worksheet);
                SetWidths(worksheet);
                SetHeights(worksheet);

                package.Save();
            }
        }
        private void AddTitles(ExcelWorksheet worksheet)
        {
            CreateTitle(worksheet, "Stage 2 Key Inputs", 2, 1);
            CreateTitle(worksheet, "Gate 2 Key Deliverables", 2, 2);
            CreateTitle(worksheet, "SCL Rep", 2, 3);
            CreateTitle(worksheet, "SCL Forecast Start", 2, 4);
            CreateTitle(worksheet, "SCL Forecast Finish", 2, 5);
            CreateTitle(worksheet, "Due Date", 2, 6);
        }

        private void CreateTitle(ExcelWorksheet worksheet, string value, int row, int column)
        {
            worksheet.Cells[row, column].Value = value;
            worksheet.Cells[row, column, row + 1, column].Merge = true;
            worksheet.Cells[row, column, row + 1, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[row, column, row + 1, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            worksheet.Cells[row, column].Style.WrapText = true;
            worksheet.Cells[row, column].Style.Font.Bold = true;
            worksheet.Cells[row, column].Style.Font.Name = "Arial";
            worksheet.Cells[row, column].Style.Font.Size = 8;
            worksheet.Cells[row, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, column].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            worksheet.Cells[row, column, row + 1, column].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
        }
        private void SetWidths(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 33;
            worksheet.Column(2).Width = 41;
            worksheet.Column(3).Width = 6;
            worksheet.Column(4).Width = 12;
            worksheet.Column(5).Width = 12;
            worksheet.Column(6).Width = 12;
        }
        private void SetHeights(ExcelWorksheet worksheet)
        {
            worksheet.Row(2).Height = 26.14;
        }
    }
}
