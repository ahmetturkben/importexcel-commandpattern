using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace ImportExcel.Service.Interfaces
{
    public interface ISqlQueryCommand : ICommand
    {
        DataTable FillDataTable(DataTable dt);
    }
}
