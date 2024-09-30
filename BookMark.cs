using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordPartition
{
    internal class BookMark
    {
        public static DataTable SetdtBookMark()
        {
            DataTable dtBookMark = new DataTable();

            DataColumn col1 = new DataColumn();
            col1.DataType = Type.GetType("System.Int32");
            col1.ColumnName = "BookMarkNum";
            dtBookMark.Columns.Add(col1);

            DataColumn col2 = new DataColumn();
            col2.DataType = Type.GetType("System.Int32");
            col2.ColumnName = "BookMarkStart";
            dtBookMark.Columns.Add(col2);

            DataColumn col3 = new DataColumn();
            col3.DataType = Type.GetType("System.Int32");
            col3.ColumnName = "BookMarkEnd";
            dtBookMark.Columns.Add(col3);

            DataColumn col4 = new DataColumn();
            col4.DataType = Type.GetType("System.String");
            col4.ColumnName = "StartString";
            dtBookMark.Columns.Add(col4);

            DataColumn col5 = new DataColumn();
            col5.DataType = Type.GetType("System.String");
            col5.ColumnName = "EndString";
            dtBookMark.Columns.Add(col5);

            DataColumn col6 = new DataColumn();
            col6.DataType = Type.GetType("System.Int32");
            col6.ColumnName = "Count";
            dtBookMark.Columns.Add(col6);

            return dtBookMark;
        }
    }
}
