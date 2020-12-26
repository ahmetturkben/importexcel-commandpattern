using ImportExcel.BLL.Models;
using ImportExcel.Service.Interfaces;
using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ImportExcel.Service.Command
{
    public class OrderCommand : IOrderCommand
    {
        private SqlQueryCommand _sqlQueryCommand;
        public OrderCommand(SqlQueryCommand sqlQueryCommand)
        {
            this._sqlQueryCommand = sqlQueryCommand;
        }

        public ServiceResult Execute()
        {
            ServiceResult sr = new ServiceResult();
            try
            {
                this._sqlQueryCommand.Execute();
                sr.IsSuccess = true;
            }
            catch
            {
                sr.IsSuccess = false;
            }

            Console.WriteLine("Order data imported succesfull!");

            return sr;
        }
    }
}
