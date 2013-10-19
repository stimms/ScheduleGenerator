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
                AddDynamicTitles();

                SetWidths();
                SetHeights();

                AddKeyInputs();

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

        private void AddDynamicTitles()
        {
            int column = 7;
            var startDate = _gates.Where(x => x.DueDate.HasValue).Select(x => x.DueDate).Min().Value.GetMonday();
            var endDate = _gates.Where(x => x.DueDate.HasValue).Select(x => x.DueDate).Max().Value.GetMonday();
            var week = startDate;
            while(week<endDate)
            {
                _worksheet.Cells[2, column++].Value = week.ToString("dd-MMM-yyyy");
                week = week.AddDays(7);
            }
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
        private void AddKeyInputs()
        {
            int row = 4;
            foreach (var majorGate in _gates.Where(x => x.Order == x.SubOrder).OrderBy(x => x.Order))
            {
                _worksheet.Cells[row, 1].Value = String.Format("{0}.  {1}", majorGate.Order, majorGate.MajorGate);

                foreach (var gate in _gates.Where(x => x.SubOrder >= Math.Floor(majorGate.Order) && x.SubOrder < Math.Floor(majorGate.Order + 1)).OrderBy(x => x.SubOrder))
                {
                    _worksheet.Cells[row, 2].Value = gate.Deliverable;
                    _worksheet.Cells[row, 2].Style.Indent = GetNumberOfDecimals(gate.SubOrder);
                    _worksheet.Cells[row, 3].Value = gate.SCLRep;
                    _worksheet.Cells[row, 4].Value = gate.SCLStartDate.HasValue ? gate.SCLStartDate.Value.ToAppliationDate() : "";
                    _worksheet.Cells[row, 5].Value = gate.SCLEndDate.HasValue ? gate.SCLEndDate.Value.ToAppliationDate() : "";
                    _worksheet.Cells[row, 6].Value = gate.DueDate.HasValue ? gate.DueDate.Value.ToAppliationDate() : "";
                    row++;
                }
                row++;
            }
        }

        private int GetNumberOfDecimals(decimal number)
        {
            return number.ToString().Split('.').Last().Length;
        }
    }
}
