using ImportExcel.Service;
using System;
using ImportExcel.Service.Command;

namespace ImportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Invoker invoiker = new Invoker();

            invoiker.SetOnStart(new OrderCommand(new SqlQueryCommand(new ExcelCommand())));
            invoiker.Running();
        }
    }
}
