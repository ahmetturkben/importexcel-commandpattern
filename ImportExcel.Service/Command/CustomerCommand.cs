using ImportExcel.BLL.Models;
using ImportExcel.Service.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ImportExcel.Service.Command
{
    public class CustomerCommand : ICommand
    {
        public ServiceResult Execute()
        {
            Console.WriteLine("Customer data imported successfull!");
            return new ServiceResult();
        }
    }
}
