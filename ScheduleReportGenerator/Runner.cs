using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleReportGenerator
{
    class Runner
    {
        private FileInfo _outputFileName;
        public void RunAsync(Action onComplete = null)
        {
            var buildSheet = new Task(() => RunGeneration());
            var openFile = buildSheet.ContinueWith(_ =>
                {
                    ExcelLauncher.DisplaySheet(_outputFileName);
                });
            openFile.ContinueWith(_ =>
            {
                if (onComplete != null)
                    onComplete();
            });
            buildSheet.Start();
        }

        private void RunGeneration()
        {
            _outputFileName = GetFileName();
            var generator = new Generator();
            generator.Generate(_outputFileName);
        }
        private FileInfo GetFileName()
        {
            return new FileInfo(Path.Combine(Environment.CurrentDirectory, String.Format("Schedule {0} {1}.xlsx", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString().Replace(":", "-"))));
        }
    }
}
