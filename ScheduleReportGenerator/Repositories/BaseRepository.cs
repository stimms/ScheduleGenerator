using System;
using Dapper;
using System.Linq;
using System.Collections.Generic;
using ScheduleReportGenerator.Models;
using System.Data;

namespace ScheduleReportGenerator.Repositories
{
    class BaseRepository
    {
        protected System.Data.IDbConnection GetConnection()
        {
            IDbConnection connection;
            try
            {
                connection = new System.Data.OleDb.OleDbConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;", AppSettings.DatabasePath));
                connection.Open();
            }
            catch(Exception ex)
            {
                connection = new System.Data.Odbc.OdbcConnection("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + AppSettings.DatabasePath);
                connection.Open();
            }
            return connection;
        }
    }
}
