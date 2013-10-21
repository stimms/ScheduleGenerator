using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using ScheduleReportGenerator.Models;
using ScheduleReportGenerator.Extensions;
using ScheduleReportGenerator.Repositories;

namespace ScheduleReportGenerator
{
    class Generator
    {
        private ExcelWorksheet _worksheet;
        private IEnumerable<Gate> _gates;

        public void Generate(FileInfo fileName)
        {
            using (var package = new ExcelPackage(fileName))
            {
                _worksheet = package.Workbook.Worksheets.Add("Schedule");
                _worksheet.Cells[1, 1, 500, 500].Style.Font.Name = "Arial";
                _worksheet.Cells[1, 1, 500, 500].Style.Font.Size = 8;
                _gates = new GateRepository().GetGates();

                AddStaticTitles();
                var columnMap = AddDynamicTitles();

                SetWidths();
                SetHeights();

                AddKeyInputs(columnMap);

                package.Save();
            }
        }
        private void AddStaticTitles()
        {
            CreateTitle("Stage 2 Key Inputs", 2, 1);
            CreateTitle("Gate 2 Key Deliverables", 2, 2);
            CreateTitle("SCL Rep", 2, 3);
            CreateTitle("SCL Forecast Start", 2, 4);
            CreateTitle("SCL Forecast Finish", 2, 5);
            CreateTitle("Due Date", 2, 6);
        }

        private IEnumerable<KeyValuePair<DateTime, int>> AddDynamicTitles()
        {
            int column = 7;
            int groupStartColumn = column;
            bool fillAlternate = false;

            var startDate = _gates.Where(x => x.DueDate.HasValue).Select(x => x.DueDate).Min().Value.GetMonday();
            var endDate = _gates.Where(x => x.DueDate.HasValue).Select(x => x.DueDate).Max().Value.GetMonday().AddDays(7);
            var week = startDate;
            var groupStartDate = week;

            var columnMap = new List<KeyValuePair<DateTime, int>>();

            while (week < endDate)
            {
                _worksheet.Cells[3, column].Value = week.ToString("dd-MMM");
                columnMap.Add(new KeyValuePair<DateTime, int>(week, column));
                week = week.AddDays(7);
                if (week.Month != groupStartDate.Month)
                {
                    _worksheet.Cells[2, groupStartColumn, 2, column].Merge = true;
                    _worksheet.Cells[2, groupStartColumn].Value = groupStartDate.ToString("MMM");
                    _worksheet.Cells[2, groupStartColumn].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    _worksheet.Cells[2, groupStartColumn].Style.Font.Bold = true;
                    _worksheet.Cells[2, groupStartColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    _worksheet.Cells[2, groupStartColumn].Style.Fill.BackgroundColor.SetColor(fillAlternate ? System.Drawing.Color.FromArgb(196, 215, 155) : System.Drawing.Color.FromArgb(184, 204, 228));
                    _worksheet.Cells[2, groupStartColumn, 2, column].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    fillAlternate = !fillAlternate;
                    groupStartDate = week;
                    groupStartColumn = column + 1;
                }
                column++;
            }
            return columnMap;
        }
        private void CreateTitle(string value, int row, int column)
        {
            _worksheet.Cells[row, column].Value = value;
            _worksheet.Cells[row, column, row + 1, column].Merge = true;
            _worksheet.Cells[row, column, row + 1, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            _worksheet.Cells[row, column, row + 1, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            _worksheet.Cells[row, column].Style.WrapText = true;
            _worksheet.Cells[row, column].Style.Font.Bold = true;
            _worksheet.Cells[row, column].Style.Font.Name = "Arial";
            _worksheet.Cells[row, column].Style.Font.Size = 8;
            _worksheet.Cells[row, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            _worksheet.Cells[row, column].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            _worksheet.Cells[row, column, row + 1, column].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
        }

        private void SetWidths()
        {
            _worksheet.Column(1).Width = 33;
            _worksheet.Column(2).Width = 41;
            _worksheet.Column(3).Width = 6;
            _worksheet.Column(4).Width = 12;
            _worksheet.Column(5).Width = 12;
            _worksheet.Column(6).Width = 12;
        }
        private void SetHeights()
        {
            _worksheet.Row(2).Height = 26.14;
        }
        private void AddKeyInputs(IEnumerable<KeyValuePair<System.DateTime, System.Int32>> columnMap)
        {
            int row = 4;
            foreach (var majorGate in _gates.GroupBy(g=>g.Order).OrderBy(x => x.Key))
            {
                _worksheet.Cells[row, 1].Value = String.Format("{0}.  {1}", majorGate.First().Order, majorGate.First().MajorGate);
                var groupStartRow = row;
                foreach (var gate in _gates.Where(x => x.SubOrder >= Math.Floor(majorGate.First().Order) && x.SubOrder < Math.Floor(majorGate.First().Order + 1)).OrderBy(x => x.SubOrder))
                {
                    _worksheet.Cells[row, 2].Value = gate.Deliverable;
                    _worksheet.Cells[row, 2].Style.Indent = GetNumberOfDecimals(gate.SubOrder);
                    _worksheet.Cells[row, 3].Value = gate.SCLRep;
                    _worksheet.Cells[row, 4].Value = gate.SCLStartDate.HasValue ? gate.SCLStartDate.Value.ToAppliationDate() : "";
                    _worksheet.Cells[row, 5].Value = gate.SCLEndDate.HasValue ? gate.SCLEndDate.Value.ToAppliationDate() : "";
                    _worksheet.Cells[row, 6].Value = gate.DueDate.HasValue ? gate.DueDate.Value.ToAppliationDate() : "";
                    if (gate.DueDate.HasValue)
                    {
                        var cell = _worksheet.Cells[row, columnMap.Where(x => x.Key == gate.DueDate.Value.GetMonday()).First().Value];
                        cell.Value = gate.Deliverable;
                        cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        cell.Style.Font.Bold = gate.Special;
                    }
                    row++;
                }
                _worksheet.Cells[groupStartRow, 1, row, columnMap.Select(x => x.Value).Max()].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                row++;

                _worksheet.Cells[1, 1, row, 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 2, row, 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 3, row, 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 4, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 5, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 6, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                _worksheet.Cells[1, 7, row, 7].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            }
        }

        private int GetNumberOfDecimals(decimal number)
        {
            return number.ToString().Split('.').Last().Length;
        }
    }
}
