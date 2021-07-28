using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using ClosedXML.Excel;

namespace Sql2Xlsx
{
    public class Oracle
    {
        public static void Write2Xlsx(Field[][] rows, string OutputPath, string SheetName)
        {
            if (rows.Length == 0)
            {
                return;
            }
            bool exists = File.Exists(OutputPath);
            var wb = exists ? new XLWorkbook(OutputPath) : new XLWorkbook();
            var ws = wb.Worksheets.Add(SheetName);

            for (int i = 0; i < rows[0].Length; i++)
            {
                ws.Cell(1, i + 1).Value = rows[0][i].Name;
            }

            for (int row = 0; row < rows.Length; row++)
            {
                for (int i = 0; i < rows[row].Length; i++)
                {
                    ws.Cell(row + 2, i + 1).Value = rows[row][i].ObjValue;
                }
            }

            // the range for which you want to add a table style
            var range = ws.Range(1, 1, rows.Length + 1, rows[0].Length);

            // create the actual table
            var table = range.CreateTable();

            // apply style
            table.Theme = XLTableTheme.TableStyleLight12;
            if (exists) wb.Save(); else wb.SaveAs(OutputPath);
        }
        public static void Write2Text(Field[][] rows, string OutputPath)
        {
            using (StreamWriter sw = new StreamWriter(OutputPath, false))
            {
                foreach (Field[] fields in rows)
                {
                    if (fields.Length == 0)
                    {
                        Console.WriteLine("Found an empty row");
                        continue;
                    }
                    sw.WriteLine("Row " + fields[0].RowCount + ":");
                    foreach (Field field in fields)
                    {
                        sw.Write(field.FieldCount + " - " + field.Name + " of " + field.DataType + ": " + field.ObjValue.ToString() + ";");
                    }
                    sw.WriteLine("");
                }
            }
        }
        public static string ReadSql(string path)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                return sr.ReadToEnd();
            }
        }
        public static async Task<string> ReadSqlAsync(string path)
        {
            using (StreamReader sr = new StreamReader(path))
            {
                return await sr.ReadToEndAsync();
            }
        }
        public static IEnumerable<Field[]> Read(string ExtractSql)
        {
            var connStr = ConfigurationManager.ConnectionStrings["EndurDB"].ConnectionString;
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                string viewschema = ConfigurationManager.AppSettings["ViewSchema"];
                using (OracleCommand cmd = new OracleCommand("alter session set CURRENT_SCHEMA = " + viewschema, conn))
                {
                    int res = cmd.ExecuteNonQuery(); // -1 is ok
                    //Console.WriteLine("session set result: " + res);
                }
                using (OracleCommand cmd = new OracleCommand(ExtractSql, conn))
                {
                    OracleDataReader DR = cmd.ExecuteReader();
                    int count = 0;
                    while (DR.Read())
                    {
                        count++;
                        Field[] fields = new Field[DR.FieldCount];
                        for (int i = 0; i < DR.FieldCount; i++)
                        {
                            Field field = new Field();
                            field.RowCount = count;
                            field.FieldCount = i;
                            field.Name = DR.GetName(i);
                            field.DataType = DR.GetDataTypeName(i);
                            // decimal overflow workaround 
                            field.ObjValue =
                                field.DataType == "Decimal"?
                                (
                                    DR.IsDBNull(i) ?
                                    (double?)null :
                                    DR.GetDouble(i)
                                ) 
                                :
                                DR.GetValue(i);
                            fields[i] = field;
                        }
                        yield return fields;
                    }
                }
            }
        }
        // private because it does not make sense to use it if it only falls back to the sync version
        private static async Task<Field[][]> ReadAsync(string ExtractSql)
        {
            var connStr = ConfigurationManager.ConnectionStrings["EndurDB"].ConnectionString;
            // https://stackoverflow.com/questions/63907271/is-oracle-openasync-etc-not-a-truly-async-method
            List<Field[]> ret = new List<Field[]>();
            await Task.Run(async () =>
            {
                using (OracleConnection conn = new OracleConnection(connStr))
                {
                    await conn.OpenAsync();
                    string viewschema = ConfigurationManager.AppSettings["ViewSchema"];
                    using (OracleCommand cmd = new OracleCommand("alter session set CURRENT_SCHEMA = " + viewschema, conn))
                    {
                        int res = await cmd.ExecuteNonQueryAsync(); // -1 is ok
                                                                    //Console.WriteLine("session set result: " + res);
                    }
                    using (OracleCommand cmd = new OracleCommand(ExtractSql, conn))
                    {
                        // https://stackoverflow.com/questions/63907271/is-oracle-openasync-etc-not-a-truly-async-method
                        OracleDataReader DR = (OracleDataReader) await cmd.ExecuteReaderAsync();
                        int count = 0;

                        while (await DR.ReadAsync())
                        {
                            count++;
                            Field[] fields = new Field[DR.FieldCount];
                            for (int i = 0; i < DR.FieldCount; i++)
                            {
                                Field field = new Field();
                                field.RowCount = count;
                                field.FieldCount = i;
                                field.Name = DR.GetName(i);
                                field.DataType = DR.GetDataTypeName(i);
                                field.ObjValue =
                                    field.DataType == "Decimal" ?
                                    DR.GetDouble(i) :
                                    DR.GetValue(i);
                                fields[i] = field;
                            }
                            ret.Add(fields);
                        }
                    }
                }

            });
            return ret.ToArray();
        }
    }
}
