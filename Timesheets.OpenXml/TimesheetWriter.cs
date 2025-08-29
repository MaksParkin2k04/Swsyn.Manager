using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Timesheets.OpenXml {
    public class TimesheetWriter {

        public static void Write(string writePath, Timesheet timesheet, string contractor, string customer) {
            WriteTemplate(writePath);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(writePath, true)) {
                WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                SharedStringTablePart shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();

                // Записываем имя исполнителя
                WriteContractor(contractor, sheetData, shareStringPart);
                // Записываем название заказчика
                WriteCustomer(customer, sheetData, shareStringPart);
                // Записываем дату начала недели
                WriteTableDate(timesheet.StartDate, sheetData);
                // Обновляем даты дней недели
                UpdateDayDate(timesheet.StartDate, sheetData);

                uint rowIndex = 10;

                foreach (ProjectTask taskRow in timesheet.Tasks) {
                    WriteTaskRow(taskRow, rowIndex, sheetData, shareStringPart, mergeCells);
                    rowIndex++;
                }

                WriteIntermediateResultsRow(timesheet, rowIndex++, sheetData, shareStringPart, mergeCells);
                WriteResultRow(timesheet, rowIndex++, sheetData, shareStringPart, mergeCells);

                for (int i = 0; i < 6; i++) {
                    WriteEmptyRow(rowIndex, sheetData);
                    rowIndex++;
                }

                for (int i = 0; i < 9; i++) {
                    WriteEmptyNanRow(rowIndex, sheetData);
                    rowIndex++;
                }

                for (int i = 0; i < 300; i++) {
                    WriteEmptyNanRow2(rowIndex, sheetData);
                    rowIndex++;
                }

                worksheetPart.Worksheet.SheetDimension = new SheetDimension() { Reference = $"A1:Y{rowIndex}" };
            }
        }

        /// <summary>
        /// Копирует шаблон таблицы Excel по указанному адресу
        /// </summary>
        /// <param name="filePath">Путь к конечному файлу</param>
        private static void WriteTemplate(string filePath) {
            string? folderPath = System.IO.Path.GetDirectoryName(filePath);
            if (Directory.Exists(folderPath) == false) {
                Directory.CreateDirectory(folderPath);
            }
            File.Copy("Templates/Template.xlsx", filePath, true);
        }

        private static void WriteContractor(string contractor, SheetData sheetData, SharedStringTablePart shareStringPart) {
            Row? row = GetRow(sheetData, 3);
            Cell? cell = GetCell(row, "C");

            int index = InsertSharedStringItem(contractor, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        private static void WriteCustomer(string customer, SheetData sheetData, SharedStringTablePart shareStringPart) {
            Row? row = GetRow(sheetData, 4);
            Cell? cell = GetCell(row, "C");

            int index = InsertSharedStringItem(customer, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        /// <summary>
        /// Обновляет дату начала таблицы в шаблоне
        /// </summary>
        /// <param name="date">Дата начала</param>
        /// <param name="sheetData">Объект листа Excel</param>
        private static void WriteTableDate(DateTime date, SheetData sheetData) {
            Cell? cell = GetCell(sheetData, "L", 5);
            double value = date.ToOADate();
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Обновляет даты дней в шаблоне таблицы
        /// </summary>
        /// <param name="startDate">Дата начала</param>
        /// <param name="sheetData">Объект листа Excel</param>
        private static void UpdateDayDate(DateTime startDate, SheetData sheetData) {
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            for (int i = 0; i < 7; i++) {
                Cell? cell = GetCell(sheetData, dateCells[i], 8);
                DateTime dateTime = startDate.AddDays(i);
                double value = dateTime.ToOADate();
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static void WriteTaskRow(ProjectTask taskRow, uint rowIndex, SheetData sheetData, SharedStringTablePart shareStringPart, MergeCells mergeCells) {
            string[] cellNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [29, 31, 27, 27, 13, 13, 13, 13, 13, 13, 13, 8, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };


            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["1:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);

            // Объединяем ячейки CD в строке
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"C{rowIndex}:D{rowIndex}") });

            WriteTypeService(taskRow.TypeService, row, shareStringPart);
            WriteTask(taskRow.Task, row, shareStringPart);
            // Записываем значение Billable в ячейку соответствующую дате
            for (int i = 0; i < taskRow.TimeDays.Count; i++) {
                TimeDay timeDay = taskRow.TimeDays[i];
                Cell? dateCell = GetCell(row, dateCells[i]);
                WriteDoubleValue(dateCell, timeDay.Billable);
            }

            // Записываем значение в вычисляемую ячейку строки
            Cell? sumCell = GetCell(row, "L");
            sumCell.CellFormula = new CellFormula($"SUM(E{rowIndex}:K{rowIndex})");
            double? sum = taskRow.GetSumBillable();
            WriteDoubleValue(sumCell, sum);
        }
        private static void WriteIntermediateResultsRow(Timesheet timesheet, uint rowIndex, SheetData sheetData, SharedStringTablePart shareStringPart, MergeCells mergeCells) {
            string[] cellNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [29, 35, 36, 37, 38, 38, 38, 38, 38, 38, 38, 10, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["1:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);

            // Объединяем ячейки BCD в строке
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"B{rowIndex}:D{rowIndex}") });
            // Записываем заголовок
            Cell titleCell = GetCell(row, "B");
            WriteSharedString("ПРОМЕЖУТОЧНЫЕ ИТОГИ", titleCell, shareStringPart);

            for (int i = 0; i < 7; i++) {
                DateTime dateTime = timesheet.StartDate.AddDays(i);
                double? sumBillable = timesheet.GetSumBillableFromDay(dateTime);

                Cell? cell = GetCell(row, dateCells[i]);
                cell.CellFormula = new CellFormula($"SUM({dateCells[i]}10:{dateCells[i]}{10 + timesheet.Tasks.Count - 1})");
                WriteDoubleValue(cell, sumBillable);
            }

            // Записываем значение в вычисляемую ячейку строки
            Cell? sumAllCell = GetCell(row, "L");
            sumAllCell.CellFormula = new CellFormula($"SUM(E{rowIndex}:K{rowIndex})");
            double? sumAll = timesheet.GetSumBillableAll();
            WriteDoubleValue(sumAllCell, sumAll);
        }
        private static void WriteResultRow(Timesheet timesheet, uint rowIndex, SheetData sheetData, SharedStringTablePart shareStringPart, MergeCells mergeCells) {
            string[] cellNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [29, 32, 33, 34, 9, 9, 9, 9, 9, 9, 9, 10, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["1:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);

            // Объединяем ячейки BCD в строке
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"B{rowIndex}:D{rowIndex}") });
            // Записываем заголовок
            Cell titleCell = GetCell(row, "B");
            WriteSharedString("ИТОГО ЧАСОВ", titleCell, shareStringPart);

            for (int i = 0; i < 7; i++) {
                DateTime dateTime = timesheet.StartDate.AddDays(i);
                double? sumBillable = timesheet.GetSumBillableFromDay(dateTime);

                Cell? cell = GetCell(row, dateCells[i]);
                cell.CellFormula = new CellFormula($"SUM({dateCells[i]}10:{dateCells[i]}{10 + timesheet.Tasks.Count - 1})");
                WriteDoubleValue(cell, sumBillable);
            }

            // Записываем значение в вычисляемую ячейку строки
            Cell? sumAllCell = GetCell(row, "L");
            sumAllCell.CellFormula = new CellFormula($"SUM(E{rowIndex}:K{rowIndex})");
            double? sumAll = timesheet.GetSumBillableAll();
            WriteDoubleValue(sumAllCell, sumAll);
        }
        private static void WriteEmptyRow(uint rowIndex, SheetData sheetData) {
            string[] cellNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [28, 28, 12, 12, 2, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["1:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);
        }
        private static void WriteEmptyNanRow(uint rowIndex, SheetData sheetData) {
            string[] cellNames = ["E", "F", "G", "H", "I", "J", "K", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [11, 11, 11, 11, 11, 11, 11, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["1:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        private static void WriteEmptyNanRow2(uint rowIndex, SheetData sheetData) {
            string[] cellNames = ["E", "F", "G", "H", "I", "J", "K", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"];
            uint[] cellStyles = [11, 11, 11, 11, 11, 11, 11, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1];
            // Название ячеек определенного дня недели
            string[] dateCells = new string[] { "E", "F", "G", "H", "I", "J", "K" };

            Row row = new Row() {
                RowIndex = rowIndex,
                Spans = new ListValue<StringValue>(["5:25"]),
                Height = 21,
                CustomHeight = true,
                DyDescent = 0.3
            };

            for (int i = 0; i < cellNames.Length; i++) {
                Cell cell = new Cell() { CellReference = $"{cellNames[i]}{rowIndex}", StyleIndex = cellStyles[i] };
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        private static void WriteTypeService(string typeService, Row row, SharedStringTablePart shareStringPart) {
            Cell? cell = GetCell(row, "B");

            int index = InsertSharedStringItem(typeService, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        private static void WriteTask(string task, Row row, SharedStringTablePart shareStringPart) {
            Cell? cell = GetCell(row, "C");

            int index = InsertSharedStringItem(task, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }

        private static Row? GetRow(SheetData sheetData, uint rowIndex) {
            return sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex == rowIndex);
        }
        private static Cell? GetCell(Row? row, string columnName) {
            if (row == null) { return null; }
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0);
        }
        private static Cell? GetCell(SheetData sheetData, string columnName, uint rowIndex) {
            Row? row = GetRow(sheetData, rowIndex);
            if (row == null) { return null; }
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
            // If the part does not contain a SharedStringTable, create one.
            shareStringPart.SharedStringTable ??= new SharedStringTable();

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>()) {
                if (item.InnerText == text) {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));

            return i;
        }
        private static void WriteSharedString(string text, Cell cell, SharedStringTablePart shareStringPart) {
            int index = InsertSharedStringItem(text, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        private static void WriteDoubleValue(Cell? cell, double? value) {
            // Устанавливаем тип данных по умолчанию (null ==> CellValues.Number)
            cell.DataType = null;
            if (value != null) {
                cell.CellValue = new CellValue(value.Value.ToString(CultureInfo.InvariantCulture));
            } else {
                cell.CellValue = null;
            }
        }
    }
}
