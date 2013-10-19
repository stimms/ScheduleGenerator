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

        public static DateTime GetMonday(this DateTime date)
        {
            var outputDate = date;
            while (outputDate.DayOfWeek != DayOfWeek.Monday)
                outputDate = outputDate.AddDays(-1);
            return outputDate;
        }
    }
}
