using System;
using Dapper;
using System.Linq;
using System.Collections.Generic;
using ScheduleReportGenerator.Models;

namespace ScheduleReportGenerator.Repositories
{
    class BaseRepository
    {
        public System.Data.IDbConnection GetConnection()
        {
            var connection = new System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}", AppSettings.DatabasePath));
            connection.Open();
            return connection;
        }
    }
    class GateRepository : BaseRepository
    {
        public IEnumerable<Gate> GetGates()
        {
            using (var connection = GetConnection())
            {
                return connection.Query<Gate>("select * from tblGate2");
            }
        }
    }
}
