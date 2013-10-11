using System;
using System.Linq;
using System.Collections.Generic;

namespace ScheduleReportGenerator.Models
{
    class Gate
    {
        public int Id { get; set; }
        public string Deliverable { get; set; }
        public bool Special { get; set; }
        public DateTime? DueDate { get; set; }
        public DateTime? SCLSTartDate { get; set; }
        public DateTime? SCLEndDate { get; set; }

        public decimal Order { get; set; }
        public decimal SubOrder { get; set; }
    }
}
