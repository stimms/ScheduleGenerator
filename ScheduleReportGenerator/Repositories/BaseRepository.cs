using System;
using Dapper;
using System.Linq;
using System.Collections.Generic;
using ScheduleReportGenerator.Models;

namespace ScheduleReportGenerator.Repositories
{
    class BaseRepository
    {
        protected System.Data.IDbConnection GetConnection()
        {
            var connection = new System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}", AppSettings.DatabasePath));
            connection.Open();
            return connection;
        }
    }
}
