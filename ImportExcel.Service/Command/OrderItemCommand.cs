using ImportExcel.BLL.Models;
using ImportExcel.Service.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ImportExcel.Service.Command
{
    class OrderItemCommand : ICommand
    {
        public ServiceResult Execute()
        {
            Console.WriteLine("Order Item data imported succesfull!");
            return new ServiceResult();
        }
    }
}
