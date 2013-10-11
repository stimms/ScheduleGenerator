using System;
using System.Linq;
using System.Collections.Generic;

namespace ScheduleReportGenerator.Extensions
{
    public static class DateTimeExtensions
    {
        public static string ToAppliationDate(this DateTime date)
        {
            return date.ToString("d-MMM-yy");
        }
    }
}
