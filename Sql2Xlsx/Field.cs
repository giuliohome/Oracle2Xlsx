using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sql2Xlsx
{
    public class Field
    {
        public int RowCount { get; set; }
        public int FieldCount { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public object ObjValue { get; set; }
    }
}
