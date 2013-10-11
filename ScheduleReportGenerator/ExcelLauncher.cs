using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace ScheduleReportGenerator
{
    class ExcelLauncher
    {
        public static void DisplaySheet(FileInfo fileName)
        {
            try
            {
                using (System.Diagnostics.Process excelProcess = new System.Diagnostics.Process())
                {
                    excelProcess.StartInfo.FileName = fileName.FullName;
                    excelProcess.Start();
                }
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
