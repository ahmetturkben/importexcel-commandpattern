using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace ImportExcel.Service.Interfaces
{
    public interface IExcelCommand : ICommand
    {
        DataTable ExecuteDataTable(DataTable dt, string conn); 
    }
}
