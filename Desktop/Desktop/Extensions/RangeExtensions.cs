using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

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

                // Add columns
                for (var colIndex = 1; colIndex <= columns.Count; colIndex++)
                {
                    Range cell = range.Cells[1, colIndex];

                    var columnName = cell.Value2?.ToString() ?? colIndex.ToString();
                    dataTable.Columns.Add(columnName);

                    Marshal.ReleaseComObject(cell);
                }

                // Add rows
                for (var rowIndex = 2; rowIndex <= rows.Count; rowIndex++)
                {
                    var dataRow = dataTable.NewRow();

                    for (var colIndex = 1; colIndex <= columns.Count; colIndex++)
                    {
                        Range cell = range[rowIndex, colIndex];

                        dataRow[colIndex - 1] = cell.Value2?.ToString() ?? "";

                        Marshal.ReleaseComObject(cell);
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