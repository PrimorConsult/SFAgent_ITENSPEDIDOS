using System;
using System.Data.Odbc;
using System.Collections.Generic;

namespace SFAgent.Sap
{
    public class SapConnector
    {
        private readonly string _connectionString;

        public SapConnector(string serverNode, string schema, string user, string password)
        {
            _connectionString = $"DRIVER={{HDBODBC}};SERVERNODE={serverNode};CURRENTSCHEMA={schema};UID={user};PWD={password};QUERYTIMEOUT=1000";
        }

        public List<Dictionary<string, object>> ExecuteQuery(string sql)
        {
            var results = new List<Dictionary<string, object>>();

            using (var connection = new OdbcConnection(_connectionString))
            {
                connection.Open();
                using (var cmd = new OdbcCommand(sql, connection))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var row = new Dictionary<string, object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[reader.GetName(i)] = reader.GetValue(i);
                        }
                        results.Add(row);
                    }
                }
            }

            return results;
        }
    }
}