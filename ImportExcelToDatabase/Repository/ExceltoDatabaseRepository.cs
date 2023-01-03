using Dapper;
using ImportExcelToDatabase.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace ImportExcelToDatabase.Repository
{
    public class ExceltoDatabaseRepository:IExceltoDatabaseRepository
    {
        private readonly BaseRepository _repo;
        public ExceltoDatabaseRepository(BaseRepository repo)
        {
            _repo = repo;
        }

        public void AddExcelDetails(List<ExceltoDatabase> excel)
        {
            for (int i = 0; i < excel.Count; i++)
            {
                string conn = "server=server6; database=CachDB15;User Id=User15;Password=csmpl#1234;Encrypt=false"; //ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                SqlConnection con = new SqlConnection(conn);
                string query = "  Insert into [CachDB15].[dbo].[ExceltoDatabase] (slno,name,email,mobile,specialization) Values('" + Convert.ToInt32(excel[i].slno) + "','" + excel[i].name.ToString() + "','" + excel[i].email.ToString() + "','" + excel[i].mobile.ToString() + "','" + excel[i].specialization.ToString() + "')";
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        public async Task<List<ExceltoDatabaseDetails>> Get()
        {
            List<ExceltoDatabaseDetails> EmpList = null;
            try
            {
                using (var connection = _repo.CreateConnection())
                {
                    DynamicParameters dynamicParam = new DynamicParameters();
                    dynamicParam.Add("@operationID", 101);
                    var result = await connection.QueryAsync<ExceltoDatabaseDetails>("sp_ExceltoDatabase", dynamicParam, commandType: CommandType.StoredProcedure);
                    return result.ToList();
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            return EmpList;
        }
    }
}
