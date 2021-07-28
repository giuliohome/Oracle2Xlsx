using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sql2Xlsx;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string oldsql = Oracle.ReadSql(@"C:\sviluppi\Endur\code\git\EndurAccruals\EndurAccruals\Scripts\query_old.sql");
            string newsql = Oracle.ReadSql(@"C:\sviluppi\Endur\code\git\EndurAccruals\EndurAccruals\Scripts\query_new.sql");
            Field[][] rows_old = Oracle.Read(oldsql).ToArray();
            Field[][] rows_new = Oracle.Read(newsql).ToArray();
            Field[] strange = rows_new.FirstOrDefault(r =>
            {
                object newcargo = r.First(f => f.Name == "CARGO_ID").ObjValue;
                return (newcargo != null &&
                    rows_old.FirstOrDefault(o => (double)o.First(f1 => f1.Name == "CARGO_ID").ObjValue == (double)newcargo)
                    == null);

            });

            return;
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: ConsoleTest.exe sql.txt output.xlsx sheetname");
                return;
            }
            Console.WriteLine("Querying DB");
            string ExtractSql = Oracle.ReadSql(args[0]);
            string OutputPath = args[1];
            string SheetName = args[2];
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            Field[][] rows = Oracle.Read(ExtractSql).ToArray();
            stopwatch.Stop();
            Console.WriteLine("DB extraction: count = " + rows.Length + " in " + stopwatch.ElapsedMilliseconds + "ms.");
            stopwatch.Reset();
            stopwatch.Start();
            Oracle.Write2Xlsx(rows, OutputPath, SheetName);
            stopwatch.Stop();
            Console.WriteLine("Excel written in " + stopwatch.ElapsedMilliseconds + "ms.");
            Console.WriteLine("Test concluded");
            Console.ReadKey();
        }
    }
}
