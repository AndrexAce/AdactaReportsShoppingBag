using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static AdactaInternational.AdactaReportsShoppingBag.Desktop.Services.ExcelService;
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
                values[r + 1, c] = cellValue == DBNull.Value ? "" : cellValue;
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
                table.ShowTableStyleFirstColumn = true;

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
                table.ShowTableStyleFirstColumn = true;

                // Format the values in the columns
                tableColumns = table.ListColumns;

                // Average column
                tableColumn = tableColumns["Media"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0.0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // LSD column
                tableColumn = tableColumns["LSD"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0.0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

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

        public void WriteFrequencyTableToWorksheet(Worksheet worksheet, string tableName)
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
                table.ShowTotals = true;
                table.ShowTableStyleFirstColumn = true;

                // Format the values in the columns
                tableColumns = table.ListColumns;
                // Values column
                tableColumn = tableColumns[1];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Percentage column
                tableColumn = tableColumns["Percentuale"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0.0%";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Total column
                tableColumn = tableColumns["Totale"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

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

        public void WriteCumulativeFrequencyTableToWorksheet(Worksheet worksheet, string tableName)
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
            Range? relatedTableRange = null;
            Range? relatedTableColumns = null;
            ListColumns? tableColumns = null;
            ListColumn? tableColumn = null;
            Range? tableColumnRange = null;

            try
            {
                // Find the corresponding table and position next to it (to the right with a blank column)
                var startRow = 1;
                var startCol = 1;
                tables = worksheet.ListObjects;

                if (tables.Count > 0)
                {
                    // Get the related table's position
                    table = tables[tableName];
                    relatedTableRange = table.Range;
                    relatedTableColumns = relatedTableRange.Columns;

                    // Start at the same row as the related table
                    startRow = relatedTableRange.Row;
                    // Start one column after the related table
                    startCol = relatedTableRange.Column + relatedTableColumns.Count + 1;

                    Marshal.ReleaseComObject(relatedTableColumns);
                    relatedTableColumns = null;
                    Marshal.ReleaseComObject(relatedTableRange);
                    relatedTableRange = null;
                    Marshal.ReleaseComObject(table);
                    table = null;
                }

                // Determine target range
                startCell = worksheet.Cells[startRow, startCol];
                endCell = worksheet.Cells[startRow + rowCount, startCol + colCount - 1];
                writeRange = worksheet.Range[startCell, endCell];

                // Write values in one operation
                writeRange.Value2 = values;

                // Create an Excel table from the range
                table = tables.Add(XlListObjectSourceType.xlSrcRange, writeRange,
                    XlListObjectHasHeaders: XlYesNoGuess.xlYes);
                table.Name = $"C_{tableName}";
                table.ShowTotals = true;
                table.ShowTableStyleFirstColumn = true;

                // Format the values in the columns
                tableColumns = table.ListColumns;

                // Percentage column
                tableColumn = tableColumns["Percentuale"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0.0%";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Total column
                tableColumn = tableColumns["Totale"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

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
                if (relatedTableColumns is not null) Marshal.ReleaseComObject(relatedTableColumns);
                if (relatedTableRange is not null) Marshal.ReleaseComObject(relatedTableRange);
                if (table is not null) Marshal.ReleaseComObject(table);
                if (tables is not null) Marshal.ReleaseComObject(tables);
                if (writeRange is not null) Marshal.ReleaseComObject(writeRange);
                if (endCell is not null) Marshal.ReleaseComObject(endCell);
                if (startCell is not null) Marshal.ReleaseComObject(startCell);
            }
        }

        public void WriteAdequacyTableToWorksheet(Worksheet worksheet, string tableName)
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
                table.ShowTableStyleFirstColumn = true;

                // Format the values in the columns
                tableColumns = table.ListColumns;

                // First column
                tableColumn = tableColumns["Troppo poco"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Second column
                tableColumn = tableColumns["Giusto"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Third column
                tableColumn = tableColumns["Troppo"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

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

        public void WriteIntensityTableToWorksheet(Worksheet worksheet, string tableName)
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
                table.ShowTableStyleFirstColumn = true;

                // Format the values in the columns
                tableColumns = table.ListColumns;

                // First column
                tableColumn = tableColumns["Per niente"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Second column
                tableColumn = tableColumns["Abbastanza"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Third column
                tableColumn = tableColumns["Estremamente"];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "0";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

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

        public void WriteSynopticTableToWorksheet(Worksheet worksheet, string tableName,
            SynopticTableType synopticTableType)
        {
            var rowCount = dataTable.Rows.Count;
            var colCount = dataTable.Columns.Count;

            // Build a 2D object array for bulk write
            var values = new object[rowCount, colCount];

            // Write data rows
            for (var r = 0; r < rowCount; r++)
            for (var c = 0; c < colCount; c++)
            {
                var cellValue = dataTable.Rows[r][c];
                values[r, c] = cellValue == DBNull.Value ? "" : cellValue;
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
                endCell = worksheet.Cells[startRow + rowCount - 1, colCount];
                writeRange = worksheet.Range[startCell, endCell];

                // Write values in one operation
                writeRange.Value2 = values;

                Marshal.ReleaseComObject(startCell);
                startCell = null;
                Marshal.ReleaseComObject(endCell);
                endCell = null;

                // Create an Excel table from the range
                table = tables.Add(XlListObjectSourceType.xlSrcRange, writeRange,
                    XlListObjectHasHeaders: XlYesNoGuess.xlNo);
                table.Name = tableName;
                table.ShowHeaders = false;

                // Format the values in the columns
                tableColumns = table.ListColumns;

                // Third column
                tableColumn = tableColumns[3];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                {
                    tableColumnRange.NumberFormat = "@";

                    Marshal.ReleaseComObject(tableColumnRange);
                    tableColumnRange = null;
                }

                Marshal.ReleaseComObject(tableColumn);
                tableColumn = null;

                // Fourth column
                tableColumn = tableColumns[4];
                tableColumnRange = tableColumn.DataBodyRange;
                if (tableColumnRange is not null)
                    switch (synopticTableType)
                    {
                        case SynopticTableType.Confezione:
                            tableColumnRange.NumberFormat = "0.0";
                            break;
                        case SynopticTableType.GradimentoComplessivo
                            or SynopticTableType.SoddisfazioneComplessiva
                            or SynopticTableType.PropensioneAlRiconsumo
                            or SynopticTableType.ConfrontoProdottoAbituale:
                            tableColumnRange.NumberFormat = "0.0";

                            startCell = tableColumnRange.Cells[2, 1];
                            endCell = tableColumnRange.Cells[4, 1];

                            Marshal.ReleaseComObject(tableColumnRange);
                            tableColumnRange = null;

                            tableColumnRange = worksheet.Range[startCell, endCell];
                            tableColumnRange.NumberFormat = "0%";
                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(synopticTableType), synopticTableType, null);
                    }

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

        public DataTable Transpose()
        {
            var transposed = new DataTable();

            // Add column from the first cell of the header
            transposed.Columns.Add(dataTable.Columns[0].ColumnName, typeof(object));

            // Add columns based on original rows (use row index or first column as identifier)
            for (var i = 0; i < dataTable.Rows.Count; i++)
                transposed.Columns.Add(dataTable.Rows[i][0].ToString(), typeof(object));

            // Transpose the data
            for (var i = 1; i < dataTable.Columns.Count; i++)
            {
                var newRow = transposed.NewRow();

                // Set the first column to the original column name
                newRow[0] = dataTable.Columns[i].ColumnName;

                // Transpose the row data
                for (var j = 0; j < dataTable.Rows.Count; j++) newRow[j + 1] = dataTable.Rows[j][i];

                transposed.Rows.Add(newRow);
            }

            return transposed;
        }
    }
}