using System;
using Dapper;
using System.Linq;
using System.Collections.Generic;
using ScheduleReportGenerator.Models;

namespace ScheduleReportGenerator.Repositories
{
    class GateRepository : BaseRepository
    {
        public IEnumerable<Gate> GetGates()
        {
            using (var connection = GetConnection())
            {
                return connection.Query<Gate>(@"select g.*, 
                                                       gm.Description as MajorGate,
                                                       w.Who as SCLRep 
                                                  from (tblGate2 g inner join 
                                                       tblGate2Major gm 
                                                    on g.order = gm.order) inner join 
                                                       tblwho w 
                                                    on w.id = g.whoid ");
            }
        }
    }
}
