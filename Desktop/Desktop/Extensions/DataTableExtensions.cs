using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Extensions;

internal static class DataTableExtensions
{
    extension(DataTable dataTable)
    {
        public void WriteToWorksheet(Worksheet worksheet, string tableName)
        {
            var rowCount = dataTable.Rows.Count;
            var colCount = dataTable.Columns.Count;

            // Build a 2D object array for bulk write
            var values = new object[rowCount + 1, colCount];

            // Write column headers
            for (var c = 0; c < colCount; c++) values[0, c] = dataTable.Columns[c].ColumnName;

            // Write data rows
            for (var r = 0; r < rowCount; r++)
            for (var c = 0; c < colCount; c++)
            {
                var cellValue = dataTable.Rows[r][c];
                values[r + 1, c] = cellValue.ToString() ?? "";
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
                table = tables.Add(XlListObjectSourceType.xlSrcRange, writeRange,
                    XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                table.Name = tableName;

                // Final formatting
                table.TableStyle = "TableStyleLight1";
                writeRange.Columns.AutoFit();
            }
            finally
            {
                if (table is not null) Marshal.ReleaseComObject(table);
                if (tables is not null) Marshal.ReleaseComObject(tables);
                if (writeRange is not null) Marshal.ReleaseComObject(writeRange);
                if (endCell is not null) Marshal.ReleaseComObject(endCell);
                if (startCell is not null) Marshal.ReleaseComObject(startCell);
            }
        }

        public void WriteClosedTableToWorksheet(Worksheet worksheet, string tableName)
        {
            var rowCount = dataTable.Rows.Count;
            var colCount = dataTable.Columns.Count;

            // Build a 2D object array for bulk write
            var values = new object[rowCount + 1, colCount];

            // Write column headers
            for (var c = 0; c < colCount; c++) values[0, c] = dataTable.Columns[c].ColumnName;

            // Write data rows
            for (var r = 0; r < rowCount; r++)
            for (var c = 0; c < colCount; c++)
            {
                var cellValue = dataTable.Rows[r][c];
                values[r + 1, c] = cellValue == DBNull.Value ? "" : cellValue;
            }

            Range? startCell = null;
            Range? endCell = null;
            Range? writeRange = null;
            ListObjects? tables = null;
            ListObject? table = null;
            Range? lastTableRange = null;
            Range? lastTableRows = null;
            ListColumns? tableColumns = null;
            ListColumn? tableColumn = null;
            Range? tableColumnRange = null;

            try
            {
                // Find the next available row (after last table + blank row)
                var startRow = 1;
                tables = worksheet.ListObjects;

                if (tables.Count > 0)
                {
                    // Get the last table's end row
                    table = tables[tables.Count];
                    lastTableRange = table.Range;
                    lastTableRows = lastTableRange.Rows;
                    startRow = lastTableRange.Row + lastTableRows.Count + 1;

                    Marshal.ReleaseComObject(lastTableRows);
                    lastTableRows = null;
                    Marshal.ReleaseComObject(lastTableRange);
                    lastTableRange = null;
                    Marshal.ReleaseComObject(table);
                    table = null;
                }

                // Determine target range
                startCell = worksheet.Cells[startRow, 1];
                endCell = worksheet.Cells[startRow + rowCount, colCount];
                writeRange = worksheet.Range[startCell, endCell];

                // Write values in one operation
                writeRange.Value2 = values;

                // Create an Excel table from the range
                table = tables.Add(XlListObjectSourceType.xlSrcRange, writeRange,
                    XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                table.Name = tableName;

                // Format the values in the columns
                tableColumns = table.ListColumns;
                // Values column
                tableColumn = tableColumns[tableColumns.Count - 1];
                tableColumnRange = tableColumn.DataBodyRange;
                tableColumnRange.NumberFormat = "0.0";

                Marshal.ReleaseComObject(tableColumnRange);
                tableColumnRange = null;
                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Lsd column
                tableColumn = tableColumns[tableColumns.Count];
                tableColumnRange = tableColumn.DataBodyRange;
                tableColumnRange.NumberFormat = "0.0";

                Marshal.ReleaseComObject(tableColumnRange);
                tableColumnRange = null;
                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Final formatting
                table.TableStyle = "TableStyleLight1";
                writeRange.Columns.AutoFit();
            }
            finally
            {
                if (tableColumnRange is not null) Marshal.ReleaseComObject(tableColumnRange);
                if (tableColumn is not null) Marshal.ReleaseComObject(tableColumn);
                if (tableColumns is not null) Marshal.ReleaseComObject(tableColumns);
                if (lastTableRows is not null) Marshal.ReleaseComObject(lastTableRows);
                if (lastTableRange is not null) Marshal.ReleaseComObject(lastTableRange);
                if (table is not null) Marshal.ReleaseComObject(table);
                if (tables is not null) Marshal.ReleaseComObject(tables);
                if (writeRange is not null) Marshal.ReleaseComObject(writeRange);
                if (endCell is not null) Marshal.ReleaseComObject(endCell);
                if (startCell is not null) Marshal.ReleaseComObject(startCell);
            }
        }

        public DataTable RemoveLastColumns(uint amount)
        {
            if (dataTable.Columns.Count <= amount)
                return dataTable;

            var result = dataTable.Clone();
            var columnsToRemove = dataTable.Columns.Cast<DataColumn>()
                .Skip(dataTable.Columns.Count - (int)amount)
                .ToList();

            // Remove unwanted columns
            foreach (var col in columnsToRemove)
                result.Columns.Remove(col.ColumnName);

            // Add rows to result
            foreach (DataRow row in dataTable.Rows)
                result.ImportRow(row);

            return result;
        }

        public DataTable RemoveLastRows(uint amount)
        {
            if (dataTable.Rows.Count <= amount)
                return dataTable;

            var result = dataTable.Clone();
            var rowsToKeep = dataTable.Rows.Cast<DataRow>()
                .Take(dataTable.Rows.Count - (int)amount);

            // Add rows to result
            foreach (var row in rowsToKeep)
                result.ImportRow(row);

            return result;
        }
    }
}