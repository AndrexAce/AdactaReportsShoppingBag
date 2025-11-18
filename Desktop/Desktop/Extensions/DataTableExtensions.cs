using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;

internal static class DataTableExtensions
{
    extension(System.Data.DataTable dataTable)
    {
        public void WriteToWorksheet(Worksheet worksheet, string tableName)
        {
            int rowCount = dataTable.Rows.Count;
            int colCount = dataTable.Columns.Count;

            // Build a 2D object array for bulk write
            var values = new object[rowCount + 1, colCount];

            // Write column headers
            for (int c = 0; c < colCount; c++)
            {
                values[0, c] = dataTable.Columns[c].ColumnName;
            }

            // Write data rows
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    var cellValue = dataTable.Rows[r][c];
                    values[r + 1, c] = cellValue?.ToString() ?? "";
                }
            }

            Range? startCell = null;
            Range? endCell = null;
            Range? writeRange = null;
            ListObjects? tables = null;
            ListObject? table = null;

            try
            {
                // Determine target range
                startCell = worksheet.Cells[1, 1];
                endCell = worksheet.Cells[rowCount + 1, colCount];
                writeRange = worksheet.Range[startCell, endCell];

                // Write values in one operation
                writeRange.Value2 = values;

                // Create an Excel table from the range
                tables = worksheet.ListObjects;
                table = tables.Add(XlListObjectSourceType.xlSrcRange, writeRange, XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                table.Name = tableName;

                // Final formatting
                table.TableStyle = "TableStyleLight1";
                writeRange.Columns.AutoFit();
            }
            finally
            {
                if (startCell is not null) Marshal.ReleaseComObject(startCell);
                if (endCell is not null) Marshal.ReleaseComObject(endCell);
                if (writeRange is not null) Marshal.ReleaseComObject(writeRange);
                if (table is not null) Marshal.ReleaseComObject(table);
                if (tables is not null) Marshal.ReleaseComObject(tables);
            }
        }
    }
}
