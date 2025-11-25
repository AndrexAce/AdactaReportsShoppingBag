using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;

internal static class RangeExtensions
{
    extension(Range range)
    {
        public DataTable MakeDataTable()
        {
            var dataTable = new DataTable();

            Range? columns = null;
            Range? rows = null;

            try
            {
                columns = range.Columns;
                rows = range.Rows;

                // Get all values at once
                var values = (object[,])range.Value2;

                if (values == null) return dataTable;

                var rowCount = values.GetLength(0);
                var colCount = values.GetLength(1);

                // Add columns from first row
                for (var colIndex = 1; colIndex <= colCount; colIndex++)
                {
                    var columnName = values[1, colIndex].ToString() ?? colIndex.ToString();
                    dataTable.Columns.Add(columnName);
                }

                // Add data rows (skip header row)
                for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    var dataRow = dataTable.NewRow();

                    for (var colIndex = 1; colIndex <= colCount; colIndex++)
                    {
                        var value = values[rowIndex, colIndex];
                        dataRow[colIndex - 1] = value;
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }
            finally
            {
                if (columns is not null) Marshal.ReleaseComObject(columns);
                if (rows is not null) Marshal.ReleaseComObject(rows);
            }

            return dataTable;
        }
    }
}