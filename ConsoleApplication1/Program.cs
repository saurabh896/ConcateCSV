using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            using (OleDbConnection conn = new OleDbConnection())
            {
                String path=@"D:\New folder\biostats.xlsx";
                
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path
                + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from Sheet1 "+"$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                       
                    }
                }
              foreach(DataRow row in dt.Rows)
              {
                  Console.WriteLine(row.Field<int>(0));
              }
            }
        }
    }
}
