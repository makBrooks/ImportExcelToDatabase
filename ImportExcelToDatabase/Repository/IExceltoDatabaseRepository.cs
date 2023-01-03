using ImportExcelToDatabase.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcelToDatabase.Repository
{
    public interface IExceltoDatabaseRepository
    {
        Task<List<ExceltoDatabaseDetails>> Get();
        void AddExcelDetails(List<ExceltoDatabase> excel);
    }
}
