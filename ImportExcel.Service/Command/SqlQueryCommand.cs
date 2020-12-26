using ImportExcel.BLL.Models;
using ImportExcel.Service.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Text;

namespace ImportExcel.Service.Command
{
    public class SqlQueryCommand : ISqlQueryCommand
    {
        private ExcelCommand _excelCommand;
        public SqlQueryCommand(ExcelCommand excelCommand)
        {
            this._excelCommand = excelCommand;
        }


        public ServiceResult Execute()
        {
            ServiceResult sr = new ServiceResult();
            try
            {
                DataTable dt = new DataTable();
                string conn = @"path";

                var cn = GetConnection();
                var dataTable = FillDataTable(dt);
                dataTable = WriteDataTable(dt, conn);

                using (cn)
                {
                    cn.Open();
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(cn))
                    {
                        bulkCopy.DestinationTableName = "[Order]";
                        bulkCopy.WriteToServer(dataTable);
                    }
                    cn.Close();
                }

                sr.IsSuccess = true;
                return sr;
            }
            catch
            {
                sr.IsSuccess = false;
                return sr;
            }
            
        }

        public DataTable FillDataTable(DataTable dt)
        {
            using (var adapter = new SqlDataAdapter($"SELECT TOP 0 * FROM [Order]", new SqlConnection("connectionString")))
            {
                adapter.Fill(dt);
            };

            return dt;
        }

        #region Private Methods
        private SqlConnection GetConnection()
        {
            SqlConnection cn = new SqlConnection("connectionString");

            return cn;
        }

        private DataTable WriteDataTable(DataTable dt, string conn)
        {
            dt = _excelCommand.ExecuteDataTable(dt, conn);

            return dt;
        }
        #endregion
    }
}
